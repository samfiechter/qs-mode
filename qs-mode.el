;;; qs-mode -- Spreadsheet Mode -- Tabular interface to Calc
;; Copyright (C) 2014 -- Use at 'yer own risk  -- NO WARRANTY!
;; Author: sam fiechter sam.fiechter(at)gmail
;; Version: 0.000000000000001
;; Created: 2014-03-24
;; Keywords: calc, spreadsheet
;; renamed from xlsx mode Sept-2022


;;;Code
(require 'avl-tree)
;; ;;;;;;;;;;;;; variables ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; pseudo constants
(defvar qs-range-parts-re "\\([A-Za-z]+\\)\\([0-9]+\\)\\:\\([A-Za-z]+\\)\\([0-9]+\\)")
(defvar qs-one-cell-re "$?[A-Za-z]+$?[0-9]+")
(defvar qs-range-re "$?[A-Za-z]+$?[0-9]+\\:$?[A-Za-z]+$?[0-9]+")
(defvar qs-cell-or-range-re (concat "[^A-Za-z0-9]?\\(" qs-range-re  "\\|" qs-one-cell-re  "\\)[^A-Za-z0-9\(]?"))
(defvar qs-empty-name "Quick Spreadsheet")

;;       CELL Format -- [ "A1" 0.5 "= 1/2" "%0.2g" "= 3 /2" (list of cells to calc when changes)]

;;   0 = index / cell Name
(defvar qs-c-addr 0)
;;   1 = value (formatted)
(defvar qs-c-fmtd 1)
;;   2 = value (number)
(defvar qs-c-val 2)
;;   3 = format (TBD)
(defvar qs-c-fmt 3)
;;   4 = formula
(defvar qs-c-fmla 4)
;;   5 = depends on -- list of indexes
(defvar qs-c-deps 5)

(defvar qs-lcmask (lognot (logxor (string-to-char "A") (string-to-char "a"))))

;; status vars
(defvar qs-cur-sheet nil)
(defvar qs-cur-col 0)
(defvar qs-max-col 3)
(defvar qs-col-widths (make-vector qs-max-col 5))
(defvar qs-mark-cell nil)
(defvar qs-cur-row 1)
(defvar qs-max-row 3)
(defvar qs-row-padding 5)
(defvar qs-sheets (list )) ;; list of qs-data
(defvar qs-data (avl-tree-create 'qs-avl-cmp))
(defvar qs-default-number-fmt "#,##0.00")
(defvar qs-cursor nil)

;;other
(defvar qs-input-history (list ))
(defvar qs-format-history (list ))

;; ;;;;;;;;;;;;; keymaps ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defvar qs-map  (make-sparse-keymap 'qs-map))

(define-key qs-map [left]       'qs-move-left)
(define-key qs-map [right]      'qs-move-right)
(define-key qs-map [tab]        'qs-move-right)
(define-key qs-map [up]         'qs-move-up)
(define-key qs-map [down]       'qs-move-down)
(define-key qs-map [return]     'qs-edit-cell)
(define-key qs-map [deletechar]        'qs-clear-key)
(define-key qs-map [backspace]  'qs-clear-key)
(define-key qs-map [remap set-mark-command] 'qs-set-mark)
(define-key qs-map [remap self-insert-command] 'qs-edit-cell)
(define-key qs-map (kbd "C-x f") 'qs-edit-format)
(define-key qs-map (kbd "C-x r") 'qs-resize-col-to-fit)
(define-key qs-map (kbd "C-x d") 'qs-redraw)

;;(define-key qs-map [C-R]        'qs-search-buffer)

(define-key qs-map (kbd "+")  'qs-increase-cur-col )
(define-key qs-map (kbd "-")  'qs-decrease-cur-col )



;;  ____        _           ___       _     _
;; |  _ \  __ _| |_ __ _   / / \   __| | __| |_ __
;; | | | |/ _` | __/ _` | / / _ \ / _` |/ _` | '__|
;; | |_| | (_| | || (_| |/ / ___ \ (_| | (_| | |
;; |____/ \__,_|\__\__,_/_/_/   \_\__,_|\__,_|_|    functions
;;

(defun qs-new-cell (addr) "Blank cell"
       (vector addr "" "" qs-default-number-fmt "" (list) )
       )

(defun qs-transform-fmla (from to fun)
  "transform a function from from to to moving all addresses relative to the addresses
EX:  From: A1 To: B1 Fun: = A2 / B1
     Returns: = B2 / C1

"
  (let* ((f (qs-addr-to-index from))
         (t (qs-addr-to-index to))
         (s (concat " fun "))
         (refs (qs-formula-cell-refs s)))
    (dolist (ref refs)
      (setq cr (qs-index-to-addr (+ (- (qs-addr-to-index (elt ref 0)) f) t)))
      (setq s (concat (substring s 0 (elt ref  1)) cr (substring s (elt ref 2))))
      )
    (substring s 1 -1)
    ))

(defun qs-clear-key () "delete current cell" (interactive)
       (let  (  (current-cell (concat (qs-col-letter qs-cur-col) (int-to-string qs-cur-row))))
	 (message "Deleting Cell : %s" current-cell)
         (if  (avl-tree-delete qs-data (qs-addr-to-index current-cell))
             (progn (qs-draw-all) (set-buffer-modified-p 1) )
           nil) ))

(defun qs-avl-cmp (a b)
  "This is the function used by avl tree to compare ss addresses"
  (let ((A (if (sequencep a) (elt a qs-c-addr) a)) (B (if (sequencep b) (elt b qs-c-addr) b)))  ; a or b can be vectors or addresses
    (< (qs-addr-to-index A) (qs-addr-to-index B)) ))


(defun qs-addr-to-index (a)  "Convert from ss addr (e.g. A1) to index  -- expects [A-Z]+[0-9]+"
       (let ((chra (- (string-to-char "A") 1))
             )
         (if (sequencep a)
             (progn
               (let ((out 0) (i 0))
                 (while (and  (< i (length a)) (< chra (elt a i)))
                   (setq out (+ (* 26 out)  (- (logand qs-lcmask (elt a i)) chra))) ;; -33 is mask to change case
                   (setq i (+ 1 i)))
                 (setq out (+ (* out qs-max-row) (floor (string-to-number (substring a i)))))
                 out))
           a )))

(defun qs-index-to-addr (idx)
  "Convert form ss index to addr (eg A1) -- expects integer"
  (if (integer-or-marker-p idx)
      (let* ((row (% idx qs-max-row))
             (col (- (/ (- idx row) qs-max-row) 1)))
        (concat (qs-col-letter col) (int-to-string row))
        ) "A1") )

(defun qs-col-number (column)
  "returns the index of a column id -- expects letters"
  (let ((chra (- (string-to-char "A") 1))
        (out 0) (i 0)
        (a column))
    (dotimes (j (length a))
      (aset a j (logand qs-lcmask (elt a j))) )
    (while (and  (< i (length a)) (< chra (elt a i)))
      (setq out (+ (* 26 out)  (- (elt a i) chra))) ;; -33 is mask to change case
      (setq i (+ 1 i)) )
    out
    ))

(defun qs-col-letter (a)
  "returns the letter of the column id arg -- expects int"
  (let ((out "") (n 1) (chra (string-to-char "A")))
    (while (<= 0 a)
      (setq out (concat (char-to-string (+ chra (% a 26))) out))
      (setq a  (- (/ a 26) 1)) )
    out ))

(defun qs-edit-format ( )
  "edit the format of the selected cell" (interactive)
  (let*  (  (current-cell (concat (qs-col-letter qs-cur-col) (int-to-string qs-cur-row)))
            (m (avl-tree-member qs-data (qs-addr-to-index current-cell)))
            (of (if m (elt m qs-c-fmt) qs-default-number-fmt))
            (prompt (concat "Cell " current-cell ": "))
            (nt (read-string prompt of  qs-yformat-history )))
    (if m
        (progn
          (aset m qs-c-fmla nt)
          (aset m qs-c-fmtd (qs-format-number nt (elt m qs-c-val)))
          (qs-draw-cell qs-cur-col qs-cur-row (elt m qs-c-fmtd)))
      (progn
        (setq m (qs-new-cell current-cell))
        (aset m qs-c-fmla nt)))))

(defun qs-set-mark () "set the mark to the current cell" (interactive)
       (if qs-mark-cell
           (setq qs-mark-cell nil)
         (setq qs-mark-cell (list qs-cur-col qs-cur-row)))
       (qs-move-cur-cell 0 0) )

(defun qs-edit-cell ( )
  "edit the selected cell"
  (interactive)

  (let*  (  (current-cell (concat (qs-col-letter qs-cur-col) (int-to-string qs-cur-row)))
            (prompt (concat "Cell " current-cell ": "))
            (m (avl-tree-member qs-data (qs-addr-to-index current-cell)))
            (ot (if m (if (string= "" (elt m qs-c-fmla)) (elt m qs-c-val) (elt m qs-c-fmla)) "" ))
            (nt 1) )
    (setq nt (if (equal 'return last-input-event)
                 (read-string prompt ot  qs-input-history )
               (read-string prompt (char-to-string last-input-event)   qs-input-history )))
    (qs-update-cell current-cell nt)
    (qs-draw-cell qs-cur-col qs-cur-row (qs-pad-right (elt m qs-c-fmtd) qs-cur-col))
    (qs-move-down)
    ))


(defun qs-update-cell (current-cell nt)
  "Update the value/formula  of current cell to nt"
  (let  ( (m (avl-tree-member qs-data (qs-addr-to-index current-cell)))
          )

    ;;delete this cell from its old deps
    (if m (let ((refs (qs-formula-cell-refs  (aref m qs-c-fmla))))
            (dolist (ref refs)
              (qs-del-dep (elt ref 0) current-cell))
            ) nil )
    ;; if this is a formula, eval and do deps
    (if (and (< 0 (length nt)) (= (string-to-char "=") (elt nt 0))) ; formulas start with  =
        (progn
          (if m
              (aset m qs-c-fmla nt)
            (progn
              (setq m (qs-new-cell current-cell))
              (aset m qs-c-fmla nt)
              (avl-tree-enter qs-data m)
	      (set-buffer-modified-p 1)
              ))

          (let ((s (concat nt " ")) (deps (list)))
            (setq s (replace-regexp-in-string qs-cell-or-range-re (lambda (str)
                                                                      (save-match-data
                                                                        (let ((addr (match-string 1 str)))
                                                                          (push addr deps)
                                                                          (qs-cell-val addr)))) s nil nil 1))
            (dolist (dep deps)
              (qs-add-dep dep current-cell))

            (setq nt (calc-eval (substring s 1 -1 ) ))  ;; throw away = and  ' '
            ))
      (if m (aset m qs-c-fmla "") nil))
    ;; nt is now not a formula
    (if m
        (progn
          (aset m qs-c-val nt)
          (aset m qs-c-fmtd (qs-format-number (elt m qs-c-fmt) nt)))
      (progn  ;;else
        (setq m (qs-new-cell current-cell))
        (aset m qs-c-val nt)
        (aset m qs-c-fmtd (qs-format-number qs-default-number-fmt nt))
        (avl-tree-enter qs-data m)
        ))
    ;;do eval-chain
    (dolist (a (aref m qs-c-deps))
      (if (equal a current-cell)
          nil ;; don't loop infinite
        (qs-eval-chain a current-cell)))  
    ))


(defun qs-xml-query (node child-node)
  "search an xml doc to find nodes with tags that match child-node"
  (let ((match (list (if (string= (car node) child-node) node nil)))
        (children (cddr node)))
    (if (listp children)
        (dolist (child children)
          (if (listp child)
              (setq match (append match (qs-xml-query child child-node))) nil)
          ) nil )
    (delq nil match)))

(defun qs-load (filename)
  "Load a file into the Spreadsheet."
  (interactive "FFilename:")
  (let ((patterns (list (cons "\\.xlsx$" 'qs-load-xlsx)
			(cons "\\.XLSX$" 'qs-load-xlsx)
                        (cons "\\.csv$" 'qs-load-csv)
			(cons "\\.CSV$" 'qs-load-csv)
                        ))
	(done nil)) 
    (dolist (e patterns)
      (if (not done)
	  (setq done  (if (string-match-p (car e) filename)
			  (eval (list (cdr e) filename))
			nil)
		nil )
	nil)
      )))

(defun qs-to-int (w)
  (truncate (if (numberp w) w (floor (string-to-number w)))))

(defun qs-load-xlsx (filename )
  "Try and read an XLSX file into qs-mode"
  (let ((xml nil)
        (sheets (list)) ;; filenames for sheets
        (styles (list)) ;; fn for styles
        (strings (list)) ;; fn for strings
        (sheettype "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
        (styletype "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
        (stringstype "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
        (shstrs (list) ))

    (with-temp-buffer  ;; read in the workbook file to get name/num sheets
      (erase-buffer)
      (shell-command (concat "unzip -p " filename " \"\\[Content_Types\\].xml\"") (current-buffer))
      (setq xml (xml-parse-region (buffer-end -1) (buffer-end 1)))
      (dolist (sh (qs-xml-query (car xml) "Override"))
(next-window-any-frame)        (let ((type (cdr (assoc 'ContentType (elt sh 1))))
              (name (cdr (assoc 'PartName (elt sh 1)))))
          (if (string= type sheettype) (push name sheets))
          (if (string= type styletype) (push name styles))
          (if (string= type stringstype) (push name strings))
          )))
    (setq sheets  (nreverse sheets))
    (setq styles (nreverse styles))
    (setq strings (nreverse strings))
    (dolist (sfn sheets)  ;; sfn is the sheet file name
      (dolist (stringfn strings) ;; stringfn is the string file name
        (with-temp-buffer
          (erase-buffer)
          (message "Decompressing Strings %s..." (concat "unzip -p " filename " " (substring stringfn 1)));; strings
          (shell-command (concat "unzip -p " filename " " (substring stringfn 1)) (current-buffer))
          (setq xml (xml-parse-region (buffer-end -1) (buffer-end 1)))
          (let ((ts (qs-xml-query (car xml) "t")) )
            (dolist (cell ts)
              (push (elt cell 2) shstrs)
              ))))
      (setq shstrs (vconcat  (nreverse shstrs)))

      (with-temp-buffer

        (message "Decompressing Sheet %s..." (concat "unzip -p " filename " " (substring sfn 1)) ) ;; sheet layout
        (erase-buffer)
        (shell-command (concat "unzip -p " filename " " (substring sfn 1)) (current-buffer))
        (setq xml (xml-parse-region (buffer-end -1) (buffer-end 1)))
        (let ((sheet (avl-tree-create 'qs-avl-cmp))
              (cols (qs-xml-query (car xml) "col")))
          (setq qs-data sheet)
                                        ;  -- The Cols
                                        ;

          (dolist (col cols)
            (let ((max (floor (string-to-number (cdr (assoc 'max (elt col 1))))))
                  (min (floor (string-to-number (cdr (assoc 'min (elt col 1))))))
                  (width (cdr (assoc 'width (elt col 1))))
                  (len (length qs-col-widths)) )
              (if (< len  max)
                  (setq qs-col-widths (vconcat qs-col-widths (make-vector (- max len) (qs-to-int width)))) nil )
              (dotimes (i (- max min))
                (aset qs-col-widths (+ min i) (qs-to-int width)))
              ))
          (dolist (cell (qs-xml-query (car xml) "c"))
            (let* ((range (prin1-to-string (cdr (assoc 'r (elt cell 1))) t))
                   (value (prin1-to-string (car (cddr (assoc 'v cell))) t))
                   (fmla (prin1-to-string  (car (cddr (assoc 'f cell))) t)))
              (if (string= value "nil") (setq value ""))
              (if (string= "s" (cdr (assoc 't (elt cell 1)))) (setq value (elt shstrs (floor (string-to-number value)))))
              (if (string= fmla "nil") (setq fmla ""))
              (if (string-match "\\([A-Za-z]+\\)\\([0-9]+\\)" range)
                  (let ((row (floor (string-to-number (match-string 2 range))))
                        (col (qs-col-number (match-string 1 range))))
                    (setq col (if col col 0))
                    (if (<= qs-max-row row) (setq qs-max-row (+ 5 row)))
                    (if (<= qs-max-col col)
                        (progn
                          (setq qs-col-widths (vconcat qs-col-widths (make-vector  (- col qs-max-col ) 7)))
                          (setq qs-max-col (length qs-col-widths))
                          ) nil )
                    ) nil)
              (if (string= "" fmla)
                  (qs-update-cell range value)
                  (qs-update-cell range (concat "=" fmla)))
              )))
        (push  qs-data (list qs-sheets sfn stringfn qs-col-widths)) ))
    (setq xls-cur-sheet nil)
    (qs-change-sheet 0)
    (qs-draw-all)
    ))
    
(defun qs-change-sheet (num)
  "edit the sheet number..."
 
  (if xls-cur-sheet
      (let ((savesheet (elt qs-sheets qs-cur-sheet)))
        (aset savesheet 0 qs-data)
        (aset savesheet 3 qs-col-widths)
        ) nil)
  (let ((xd (elt qs-sheets num)))
    (setq qs-data (if (car xd) (car xd) (avl-tree-create 'qs-avl-cmp)))
    (setq qs-col-widths (elt xd 3))
    ))

(defun qs-load-csv (filename)
  "Read a CSV file into qs-mode"
  (interactive "fFilename:")
    (with-temp-buffer
      (insert-file-contents filename)
      (qs-read-csv-buffer)
      ))

(defun qs-read-csv-buffer () "Read a buffer full 'o CSV"
       (let ((x 0) (cell "") (cl qs-cur-col)
             (rw qs-cur-row) (j 0)
             (re "\\(\"\\([^\"]+\\)\"\\|\\([^,]+\\)\\)[,\n]")
             (newline "")
	     (split ())
	     (lines ()) )
	 (read-only-mode 0)
	 (beginning-of-buffer)
         (while (not (eobp))
           (setq newline (thing-at-point 'line))
	   (setq split ())
           (setq j 0)
           (while (and (< j (length newline)) (string-match re newline j))
             (setq x (if (match-string 2 newline) 2 1))
             (setq cell (match-string x newline))   
             (setq j (+ 1 (match-end x)))
	     (add-to-ordered-list 'split cell (+ 1 (length split)))
	     )
	   (add-to-ordered-list 'lines split (+ 1 (length lines)))
	   (forward-line 1))
	 (erase-buffer)
	 (setq j 1)
	 (dolist (line lines)
	   (progn
	     (setq x 0)
	     (dolist (cell line)
	       (progn
		 (setq qs-cur-col x)
		 (setq qs-cur-row j)
		 (qs-update-cell (concat (qs-col-letter x) (int-to-string j)) cell)
		 (read-only-mode 0)
		 (setq x (+ x 1))
		 (if (= qs-max-col x)
                     (progn
		       (setq qs-max-col (+ x 1))
		       (setq qs-col-widths (vconcat qs-col-widths (list 5)))
		       )
		   )))
	     (setq j (+ 1 j))
	     ))
	 ;; (setq qs-cur-col 0)
	 ;;   (while (< qs-cur-col qs-max-col)
	 ;;     (qs-resize-col-to-fit)
	 ;;     (setq qs-cur-col (+ qs-cur-col 1))
	 ;;     )
	 ;;   
	 ;;   
	 (read-only-mode 1)
	 (setq qs-cur-row 1)
	 (setq qs-cur-col 0)
	 (qs-draw-all)
	   ))

(defun qs-draw-csv ()
  "Save to a CSV file"  
  (let (fn (buffer-file-name))	 
    (read-only-mode 0)
    (erase-buffer)
      (dotimes (j qs-max-row) ;; draw buffer
	(progn
	  (dotimes (i qs-max-col)
            (let ((m (avl-tree-member qs-data (qs-addr-to-index (concat (qs-col-letter i) (int-to-string (+ 1 j)))))))
	      ( if m
		  (insert (concat (if (= 0 i) "" ",") (concat "\"" (if (string= "" (elt m qs-c-fmla)) (elt m qs-c-val) (elt m qs-c-fmla))  "\"" )))
				  (insert (if (= 0 i) "" ","))
				  )))
	(insert "\n") ))
      (read-only-mode 1)
  ))

;;  ____                     _
;; |  _ \ _ __ __ ___      _(_)_ __   __ _
;; | | | | '__/ _` \ \ /\ / / | '_ \ / _` |
;; | |_| | | | (_| |\ V  V /| | | | | (_| |
;; |____/|_|  \__,_| \_/\_/ |_|_| |_|\__, |
;;                                   |___/
;; functions dealing with the cursor and cell drawing /padding



(defun qs-format-number (fmt val)
  "output text format of numerical value according to format string
   000000.00  -- pad to six digits (or however many zeros to left of .) round at two digits, or pad out to two
   #,###.0 -- insert a comma (anywhere to the left of the . is fine -- you only need one and it'll comma ever three
   0.00## -- Round to four digits, and pad out to at least two."
  (if (stringp val) nil (setq val (format "%s" val)))
  (if (equal 0 (string-match "^ *\\+?-?[0-9,\\.]+ *$" val))
      (progn
        (let ((ip 0)   ;;int padding
              (dp 0)   ;;decmel padding
              (dr 0)   ;;decmil round
              (dot-index (string-match "\\." fmt))
              (power 0)
              (value (float (if (numberp val) val (string-to-number val))))
              (dec ""))
          (dotimes (i (or dot-index (length fmt)))
            (and (= (elt fmt i) (string-to-char "0")) (setq ip (1+ ip)))
            )
          (if dot-index
              (progn (dotimes (i (- (length fmt) dot-index))
                       (and (= (elt fmt (+ i dot-index)) (string-to-char "0")) (setq dp (1+ dp)))
                       (and (= (elt fmt (+ i dot-index)) (string-to-char "#")) (setq dr (1+ dr)))
                       )
                     (setq power (+ dp dr))
                     (setq dec  (fround (* (- (+ 1 value) (ftruncate value)) (expt 10 power))))  ;; 0.0 = 1.00
                     (setq power (- (length (format "%d" dec)) dp 2))  ; -2 = "1."
                     (if (> power 0)
                         (setq dec (* dec (expt 10 power)))
                       nil
                       )
                     (setq dec (concat "." (substring (format "%d" dec) 1)))
                     ) nil )
          (setq dec (concat (format (concat "%0." (number-to-string ip) "d") (ftruncate value)) (if dot-index dec nil)))
          (if (string-match "," fmt)
              (let ((re "\\([0-9]\\)\\([0-9]\\{3\\}\\)\\([,\\.]\\)\\|\\([0-9]\\)\\([0-9][0-9][0-9]\\)$")
                    (rep (lambda (a) (save-match-data (concat (match-string 1 a) (match-string 4 a) "," (match-string 5 a) (match-string 2 a) (match-string 3 a))))))
                (while (string-match re dec)
                  (setq dec (replace-regexp-in-string re rep dec)))
                ) nil )
          dec))
    val))

(defun qs-highlight (txt)
  "highlight text"
  (let ((myface '((:foreground "White") (:background "Blue"))))
    (propertize txt 'font-lock-face myface)))

(defun qs-row-highlight (txt)
  "highlight text"
  (let ((myface '((:foreground "#2f3036") (:background "#e6eafa"))))
    (propertize txt 'font-lock-face myface)))

(defun qs-pad-center (s i)
  "pad a string out to center it - expects stirng, col no (int)"
  (setq s (replace-regexp-in-string "\n" " " s))
  (let ( (ll (length s) ))
    (if (>= ll (- (elt qs-col-widths i) 2))
        (make-string (elt qs-col-widths i) (string-to-char "#"))
      (let* ( (pl (/ (- (elt qs-col-widths i) ll)  2))  ; half the pad length
              (pad (make-string (- (elt qs-col-widths i) pl 1 ) (string-to-char " "))) )
        (concat (make-string pl (string-to-char " ")) s pad)) )
    ))

(defun qs-pad-left (s i)
  "pad to the left  - expects stirng, col no (int)"
  (setq s (replace-regexp-in-string "\n" " " s))
  (let ( (ll (length s)) )
    (if (>= ll (- (elt qs-col-widths i) 2))
        (make-string (elt qs-col-widths i) (string-to-char "#"))
      (concat s (make-string (- (elt qs-col-widths i)  ll) (string-to-char " ")))
      )))

(defun qs-pad-right (s i)
  "pad to the right  - expects stirng, col no (int)"
  (setq s (replace-regexp-in-string "\n" " " (if s s "")))
  (let ( (ll (length s)) )
    (if (>= ll (- (elt qs-col-widths i) 2))
        (make-string (elt qs-col-widths i) (string-to-char "#"))
      (concat (make-string (- (elt qs-col-widths i)  ll) (string-to-char " ")) s)
      )))


(defun qs-draw-all ()
  "Populate the current ss buffer." (interactive)
  ;;  (pop-to-buffer qs-empty-name nil)

  (let ((i 0) (j 0) (k 0) (header (make-string qs-row-padding (string-to-char " "))))
    (read-only-mode 0)
    (beginning-of-buffer)
    (erase-buffer)
    (setq cursor-type nil)  ;; no cursor
    (setq truncate-lines 1)  ;; no wrap-aroudn
    (dotimes (i qs-max-col)  ;; make header
      (setq header (concat header (qs-pad-right (qs-col-letter i) i))))
    (dotimes (j qs-max-row) ;; draw buffer
      (if (= 0 j)
          (insert  (qs-highlight header))
        (progn
          (insert (qs-highlight (format (concat "%" (int-to-string qs-row-padding) "d") j) ))
          (dotimes (i qs-max-col)
            (let ((m (avl-tree-member qs-data (+ j (* qs-max-row (+ i 1))))))
	      (let ((s (if m
			(qs-pad-right (elt m qs-c-fmtd) i)
                       (make-string (elt qs-col-widths i) (string-to-char " "))
		       )))
	      (if (= 0 (logand 1 j))
		  (insert s)
		(insert (qs-row-highlight s)))
	      )))))	      		 
      (insert "\n"))
    (set-buffer-modified-p nil)
    (qs-move-cur-cell 0 0) ;;draw cursor
    (read-only-mode 1)
    ))

(defun qs-redraw ()
  (interactive)
  (qs-draw-all)
  )

(defun qs-resize-col-to-fit ()
  (interactive)
  (let ((y 1)
	(len 3) )	
    (while (< y qs-max-row)
      (let ((m (avl-tree-member qs-data (+ y (* qs-max-row (+ 1 qs-cur-col))))))
	(if m
	    (let ((str (concat "   " (elt m qs-c-fmtd))))
	      (if (< len (length str))
		  (setq len (length str))
		nil ))
	  nil )
	(setq y (+ 1 y))
	))    
    (aset qs-col-widths qs-cur-col len)
    (qs-draw-all)   
    ) )




    

(defun qs-draw-cell (x y text)
  "redraw one cell on the ss  - expects  col no. (int), row no, (int), cell value (padded string)"
  (let ((col qs-row-padding) (i 0))
    (read-only-mode 0)
    ;;    (pop-to-buffer qs-empty-name nil)
    (setq cursor-type nil)  ;; no cursor
    (setq truncate-lines 1)  ;; no wrap-around

    (dotimes (i x) (setq col (+ col (elt qs-col-widths i))) )
    (goto-line (+ y 1))
    (move-to-column col)
    (if (= 0 (logand 1 y))
	(insert text)
      (insert (qs-row-highlight text)) )
    (delete-forward-char (length text))
    (read-only-mode 1) 
    ))

(defun qs-move-left ()
  (interactive)
  (if (> qs-cur-col  0) (qs-move-cur-cell -1 0) nil)
  )

(defun qs-move-right ()
  (interactive)
  (if (< qs-cur-col (- qs-max-col 1))
      (qs-move-cur-cell 1 0)
    (progn
      (setq qs-max-col (+ 1 qs-max-col))
      (setq qs-col-widths   (vconcat qs-col-widths (list (elt qs-col-widths qs-cur-col))))
      (setq qs-cur-col (+ 1 qs-cur-col))
      (qs-draw-all)
      )))

(defun qs-move-up ()
  (interactive)
  (if (> qs-cur-row 1)
      (qs-move-cur-cell 0 -1)
    nil ))

(defun qs-move-down ()
  (interactive)
  (if (< qs-cur-row (- qs-max-row 1))
      (qs-move-cur-cell 0 1 )
    (progn
      (setq qs-max-row (+ 1 qs-max-row))
      (setq qs-cur-row (+ 1 qs-cur-row))
      (qs-draw-all)
      )))

(defun qs-move-cur-cell (x y) (interactive)
       "Move the cursor or selection overaly"
       (goto-char 0)
       (let* ((new-row (+ qs-cur-row y))
              (new-col (+ qs-cur-col x))
              (n (avl-tree-member qs-data (+ new-row (* qs-max-row (+ 1 new-col)))))
              (row-pos (line-beginning-position (+ 1 new-row)))
              (cur-mesg "")
              (col-pos (+ 4 (if (> new-col 0) (apply '+ (mapcar (lambda (x) (elt qs-col-widths x)) (number-sequence 0 (- new-col 1)))) 0))) )
         (if (listp qs-cursor) (progn
                                   (dolist (ovl qs-cursor) (delete-overlay ovl))
                                   (setq qs-cursor nil) ))
         (if qs-mark-cell
             (progn  ;; mark set...
               (if (overlayp qs-cursor) (progn
                                            (delete-overlay qs-cursor)
                                            (setq qs-cursor nil) ))
               (let* ((mc (elt qs-mark-cell 0))
                      (mr (elt qs-mark-cell 1))
                      (minc (if (> new-col mc) mc new-col))
                      (maxc (if (> new-col mc) new-col mc))
                      (minr (if (> new-row mr) mr new-row))
                      (maxr (if (> new-row mr) new-row mr))
                      (col-min  (+ 4 (if (> minc 0) (apply '+ (mapcar (lambda (x) (elt qs-col-widths x)) (number-sequence 0 (- minc 1) ))) 0)))
                      (col-max  (+ 4 (if (>= maxc 0) (apply '+ (mapcar (lambda (x) (elt qs-col-widths x)) (number-sequence 0 maxc ))) 0))) )
                 (dotimes (row (+ 1 (- maxr minr)))
                   (let ((ovl nil) (row-pos (line-beginning-position (+ 1 row minr))))
                     (setq ovl (make-overlay (+ row-pos col-min) (+ row-pos col-max)))
                     (overlay-put ovl 'face '((:foreground "White") (:background "Blue")))
                     (push ovl qs-cursor)
                     ))
                 (setq qs-cur-row new-row)
                 (setq qs-cur-col new-col)
                 (minibuffer-message (concat "Range " (qs-col-letter mc) (int-to-string mr)
                                             ":" (qs-col-letter qs-cur-col) (int-to-string qs-cur-row)))
                 ))
           ;; Mark Not set...
           (let ((n (avl-tree-member qs-data (+ new-row (* qs-max-row (+ 1 new-col))))))
             (if (overlayp qs-cursor)
                 (move-overlay qs-cursor (+ row-pos col-pos) (+ row-pos col-pos (elt qs-col-widths new-col)))
               (progn
                 (setq qs-cursor (make-overlay (+ row-pos col-pos) (+ row-pos col-pos (elt qs-col-widths  new-col))))
                 ))

             (overlay-put qs-cursor 'face '((:foreground "White") (:background "Blue")))
             (setq qs-cur-row new-row)
             (setq qs-cur-col new-col)
             (setq cur-mesg (concat "Cell " (qs-col-letter qs-cur-col) (int-to-string qs-cur-row)
                                    " : " (if n (if (string= "" (elt n qs-c-fmla)) (elt n qs-c-val ) (elt n qs-c-fmla)) "" )))
             ))
         (goto-char (+ row-pos col-pos))
         (minibuffer-message cur-mesg)
         ))

(defun qs-increase-cur-col ()
  "Increase the width of the current column" (interactive)
  (aset qs-col-widths qs-cur-col (+ 1 (elt qs-col-widths qs-cur-col)))
  (qs-draw-all))

(defun qs-decrease-cur-col ()
  "Decrease the width of the current column" (interactive)
  (aset qs-col-widths qs-cur-col (- (elt qs-col-widths qs-cur-col) 1))
  (qs-draw-all))

;;   ____     _ _   _____            _             _   _
;;  / ___|___| | | | ____|_   ____ _| |_   _  __ _| |_(_) ___  _ __
;; | |   / _ \ | | |  _| \ \ / / _` | | | | |/ _` | __| |/ _ \| '_ \
;; | |__|  __/ | | | |___ \ V / (_| | | |_| | (_| | |_| | (_) | | | |
;;  \____\___|_|_| |_____| \_/ \__,_|_|\__,_|\__,_|\__|_|\___/|_| |_|
;; fuctions dealing with eval

(defun qs-cell-val (address)
  "get the value of a cell"
  (let ((addr (replace-regexp-in-string "\\$" "" address)))
    (if (string-match qs-range-parts-re addr)
        (progn
          (let* ((s (qs-addr-to-index (concat  (match-string 1 addr) (match-string 2 addr))))
                 (cl (- (qs-addr-to-index (concat  (match-string 1 addr) (match-string 4 addr))) s))
                 (rl (- (qs-addr-to-index (concat  (match-string 3 addr) (match-string 2 addr))) s))
                 (rv "[") (m 0)
                 )
            (if (< cl 0)
                (progn
                  (setq cl (* -1 cl))
                  (setq s (- s cl)) ) nil)
            (if (< rl 0)
                (progn
                  (setq rl (* -1 rl))
                  (setq s (- s rl))) nil )
            (setq rl (/ rl qs-max-row))
            (dotimes (r (+ 1 rl))
              (setq rv (concat rv "["))
              (dotimes (c (+ 1 cl))
                (setq m (avl-tree-member qs-data (+ s c (* qs-max-row r))))
                (setq rv (concat rv (if m (elt m qs-c-val) "0") (if (= c cl) "]" ",") ))
                )
              (setq rv (concat rv (if (= r rl) "]" ",")))
              ) rv))
      (progn
        (let ((m (avl-tree-member qs-data  (qs-addr-to-index addr))) )
          (if m (elt m qs-c-val) "0")
          )))))


(defun qs-eval-chain (addr chain)
  "Update cell and all deps"
  (let ((m (avl-tree-member qs-data (qs-addr-to-index addr))) )
    (if m
        (progn
          (qs-eval-fun addr)
          (dolist (a (elt m qs-c-deps))
            (if (member a chain) nil
              (qs-eval-chain a (append chain (list addr)))
              ))))))


(defun qs-formula-cell-refs (formula)
  "Takes a string and finds all the cell references in a string Ex:
   (qs-fomual-cell-refs (\"= A1+B1\"))
   returns a list ((\"A1\" 2 3) (\"B1\" 5 6))"
  (let* ((s (concat formula " "))
         (j 0) (retval (list )))
    (while (string-match qs-cell-or-range-re s j)
      (setq retval (append retval (list (list (match-string 1 s) (match-beginning 1) (match-end 1)))))
      (setq j (+ 1(match-end 1)))
      )
    retval))

(defun qs-eval-fun (addr)
  "sets cell value based on its function; draws"
  (let*  ( (m (avl-tree-member qs-data (qs-addr-to-index addr)))
           (cv "")
           (s (if m (elt m qs-c-fmla) ""))
           (refs (qs-formula-cell-refs s)) )
    (if refs
        (progn
          (dolist (ref refs)
            (setq cv (qs-cell-val (elt ref 0)))
            (setq s (concat (substring s 0 (elt ref  1)) cv (substring s (elt ref 2)))))
          (setq cv (calc-eval (substring s 1)))
          (aset m qs-c-val cv)
          (aset m qs-c-fmtd (qs-format-number (elt m qs-c-fmt) cv))
          (let ((c 0) (i 0) (chra (- (string-to-char "A") 1)))
            (while (and  (< i (length addr)) (< chra (elt addr i)))
              (setq c (+ (* 26 c)  (- (logand -33 (elt addr i)) chra))) ;; -33 is mask to change case
              (setq i (+ 1 i)))
            (qs-draw-cell (- c 1) (string-to-number (substring addr i))
                            (qs-pad-right (elt m qs-c-fmtd) (- c 1))) )
          cv ) (aref m qs-c-val)
          )))


(defun qs-add-dep (ca cc)
  "Add to dep list.
   CA -- addr to to add
   CC -- cell who depends"
  (let ((addr (replace-regexp-in-string "\\$" "" ca)))
    (if (string-match qs-range-parts-re addr)
        (progn
          (let* ((s (qs-addr-to-index (concat  (match-string 1 addr) (match-string 2 addr))))
                 (cl (- (qs-addr-to-index (concat  (match-string 1 addr) (match-string 4 addr))) s))
                 (rl (- (qs-addr-to-index (concat  (match-string 3 addr) (match-string 2 addr))) s))
                 )
            (if (< cl 0)
                (progn
                  (setq cl (* -1 cl))
                  (setq s (- s cl)) ) nil)
            (if (< rl 0)
                (progn
                  (setq rl (* -1 rl))
                  (setq s (- s rl))) nil )
            (setq rl (/ rl qs-max-row))
            (dotimes (r (+ 1 rl))
              (dotimes (c (+ 1 cl))
                (qs-add-dep  (qs-index-to-addr (+ s c (* r qs-max-row))) cc)
                )
              )))
      (let  ( (m (avl-tree-member qs-data (qs-addr-to-index addr))))
        (if m
            (let ((od (elt m qs-c-deps)))
              (if (member cc od) nil
                (aset m qs-c-deps (append (elt m qs-c-deps) (list cc))) ))
          (progn
            (setq m (qs-new-cell addr))
            (aset m qs-c-deps (list cc))
            (avl-tree-enter qs-data m)
            ))))))


(defun qs-del-dep (ca cc)
  "Remove from dep list."
  (let ((addr (replace-regexp-in-string "\\$" "" ca)))
    (if (string-match qs-range-parts-re addr)
        (progn
          (let* ((s (qs-addr-to-index (concat  (match-string 1 addr) (match-string 2 addr))))
                 (cl (- (qs-addr-to-index (concat  (match-string 1 addr) (match-string 4 addr))) s))
                 (rl (- (qs-addr-to-index (concat  (match-string 3 addr) (match-string 2 addr))) s))
                 )
            (if (< cl 0)
                (progn
                  (setq cl (* -1 cl))
                  (setq s (- s cl)) ) nil)
            (if (< rl 0)
                (progn
                  (setq rl (* -1 rl))
                  (setq s (- s rl))) nil )
            (setq rl (/ rl qs-max-row))
            (dotimes (r (+ 1 rl))
              (dotimes (c (+ 1 cl))
                (qs-del-dep (qs-index-to-addr (+ s c (* r qs-max-row))) cc)
                )
              )))
      (let  ( (m (avl-tree-member qs-data (qs-addr-to-index addr))))
        (if
            (delete cc (aref m qs-c-deps))
            nil) ))))



;;;###autoload
(define-derived-mode qs-mode special-mode qs-empty-name
  "ss game mode
  Keybindings:
  \\{qs-map} "
  (use-local-map qs-map)  
  ;;  (setq truncate-lines 1)

  ;; (unless (featurep 'emacs)
  ;;   (setq mode-popup-menu
  ;;         '("ss Commands"
  ;;           ["Start new game"        qs-start-game]
  ;;           ["End game"                qs-end-game
  ;;            (qs-active-p)]
  ;;           ))


  (setq qs-cur-col 0)
  (setq qs-max-col 30)
  (setq qs-col-widths (make-vector qs-max-col 15))
  (setq qs-cur-row 1)
  (setq qs-max-row 30)
  (setq qs-row-padding 4)
  (setq qs-data (avl-tree-create 'qs-avl-cmp))
  (add-hook 'before-save-hook 'qs-draw-csv 90 1)
  (add-hook 'after-save-hook 'qs-draw-all 90 1)
  (if (and (buffer-file-name) (qs-read-csv-buffer))

    nil
    )
  (qs-draw-all)
  ;;  (autoarg-kp-mode 1)


  
  )

;;  ____        __ __  __       _   _
;; |  _ \  ___ / _|  \/  | __ _| |_| |__  ___
;; | | | |/ _ \ |_| |\/| |/ _` | __| '_ \/ __|
;; | |_| |  __/  _| |  | | (_| | |_| | | \__ \
;; |____/ \___|_| |_|  |_|\__,_|\__|_| |_|___/

(defmath SUM (x)
  "add the items in the range"
  (interactive 1 "sum")
  :" vsum(x)"

  )

(defmath ROUND (x n)
  "Rounds the number x to n"
  (interactive 2 "x n")
  :" (calc-floor(x * 10^n) / 10^n)"
  )


(provide 'qs-mode)
(push (cons "\\.xlsx$" 'qs-mode) auto-mode-alist)
(push (cons "\\.XLSX$" 'qs-mode) auto-mode-alist)
(push (cons "\\.csv$" 'qs-mode) auto-mode-alist)
(push (cons "\\.CSV$" 'qs-mode) auto-mode-alist)

;;; qs-mode.el ends here

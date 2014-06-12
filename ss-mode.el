;;; ss-mode -- Spreadsheet Mode -- Tabular interface to Calc
;; Copyright (C) 2014 -- Use at 'yer own risk  -- NO WARRANTY!
;; Author: sam fiechter sam.fiechter(at)gmail
;; Version: 0.000000000000001
;; Created: 2014-03-24
;; Keywords: calc, spreadsheet


;;;Code
(require 'avl-tree)
;; ;;;;;;;;;;;;; variables ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; pseudo constants
(defvar ss-range-parts-re "\\([A-Za-z]+\\)\\([0-9]+\\)\\:\\([A-Za-z]+\\)\\([0-9]+\\)")
(defvar ss-one-cell-re "$?[A-Za-z]+$?[0-9]+")
(defvar ss-range-re "$?[A-Za-z]+$?[0-9]+\\:$?[A-Za-z]+$?[0-9]+")
(defvar ss-cell-or-range-re (concat "[^A-Za-z0-9]?\\(" ss-range-re  "\\|" ss-one-cell-re  "\\)[^A-Za-z0-9\(]?"))
(defvar ss-empty-name "*Sams Spreadsheet Mode*")

;;       CELL Format -- [ "A1" 0.5 "= 1/2" "%0.2g" "= 3 /2" (list of cells to calc when changes)]
;;   0 = index / cell Name
(defvar ss-c-addr 0)
;;   1 = value (formatted)
(defvar ss-c-fmtd 1)
;;   2 = value (number)
(defvar ss-c-val 2)
;;   3 = format (TBD)
(defvar ss-c-fmt 3)
;;   4 = formula
(defvar ss-c-fmla 4)
;;   5 = depends on -- list of indexes
(defvar ss-c-deps 5)


;; status vars
(defvar ss-cur-col 0)
(defvar ss-max-col 3)
(defvar ss-col-widths (make-vector ss-max-col 7))
(defvar ss-mark-cell nil)
(defvar ss-cur-row 1)
(defvar ss-max-row 3)
(defvar ss-row-padding 4)
(defvar ss-sheets (list )) ;; list of ss-data
(defvar ss-data (avl-tree-create 'ss-avl-cmp))
(defvar ss-default-number-fmt ",0.00")
(defvar ss-cursor nil)

;;other 
(defvar ss-input-history (list ))
(defvar ss-format-history (list ))
;; ;;;;;;;;;;;;; keymaps ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defvar ss-map  (make-sparse-keymap 'ss-map))


(define-key ss-map [left]       'ss-move-left)
(define-key ss-map [right]      'ss-move-right)
(define-key ss-map [tab]        'ss-move-right)
(define-key ss-map [up]         'ss-move-up)
(define-key ss-map [down]       'ss-move-down)
(define-key ss-map [return]     'ss-edit-cell)
(define-key ss-map [backspace]  'ss-clear-key)
(define-key ss-map [remap set-mark-command] 'ss-set-mark)
(define-key ss-map [remap self-insert-command] 'ss-edit-cell)
(define-key ss-map (kbd "C-x f") 'ss-edit-format)

;;(define-key ss-map [C-R]        'ss-search-buffer)

(define-key ss-map (kbd "+")  'ss-increase-cur-col )
(define-key ss-map (kbd "-")  'ss-decrease-cur-col )



;;  ____        _           ___       _     _
;; |  _ \  __ _| |_ __ _   / / \   __| | __| |_ __
;; | | | |/ _` | __/ _` | / / _ \ / _` |/ _` | '__|
;; | |_| | (_| | || (_| |/ / ___ \ (_| | (_| | |
;; |____/ \__,_|\__\__,_/_/_/   \_\__,_|\__,_|_|    functions
;;

(defun ss-new-cell (addr) "Blank cell"
       (vector addr "" "" ss-default-number-fmt "" (list) )
       )

(defun ss-transform-fmla (from to fun)
  "transform a function from from to to moving all addresses relative to the addresses
EX:  From: A1 To: B1 Fun: = A2 / B1
     Returns: = B2 / C1

"
  (let* ((f (ss-addr-to-index from))
         (t (ss-addr-to-index to))
         (s (concat " fun "))
         (refs (ss-formula-cell-refs s)))
    (dolist (ref refs)
      (setq cr (ss-index-to-addr (+ (- (ss-addr-to-index (elt ref 0)) f) t)))
      (setq s (concat (substring s 0 (elt ref  1)) cr (substring s (elt ref 2))))
      )
    (substring s 1 -1)
    ))

(defun ss-clear-key () "delete current cell" (interactive)
(let  (  (current-cell (concat (ss-col-letter ss-cur-col) (int-to-string ss-cur-row))))
  (if  (avl-tree-delete ss-data current-cell)
    (ss-draw-cell ss-cur-col ss-cur-row (ss-highlight (ss-pad-right "" ss-cur-col) ))
     nil) ))

(defun ss-avl-cmp (a b)
  "This is the function used by avl tree to compare ss addresses"
  (let ((A (if (sequencep a) (elt a ss-c-addr) a)) (B (if (sequencep b) (elt b ss-c-addr) b)))  ; a or b can be vectors or addresses
    (< (ss-addr-to-index A) (ss-addr-to-index B)) ))


(defun ss-addr-to-index (a)
  "Convert from ss addr (e.g. A1) to index  -- expects [A-Z]+[0-9]+"
  (let ((chra (- (string-to-char "A") 1))
        (lcmask (lognot (logxor (string-to-char "A") (string-to-char "a"))))
        )
    (if (sequencep a)
        (progn
          (let ((out 0) (i 0))
            (while (and  (< i (length a)) (< chra (elt a i)))
              (setq out (+ (* 26 out)  (- (logand lcmask (elt a i)) chra))) ;; -33 is mask to change case
              (setq i (+ 1 i)))
            (setq out (+ (* out ss-max-row) (string-to-int (substring a i))))
            out))
      a )))

(defun ss-index-to-addr (idx)
  "Convert form ss index to addr (eg A1) -- expects integer"
  (if (integer-or-marker-p idx)
      (let* ((row (% idx ss-max-row))
             (col (- (/ (- idx row) ss-max-row) 1)))
        (concat (ss-col-letter col) (int-to-string row))
        ) "A1") )

(defun ss-col-number (a)
  "returns the index of a column id -- expects letters"
  (let ((chra (- (string-to-char "A") 1))
        (lcmask (lognot (logxor (string-to-char "A") (string-to-char "a"))))
        (out 0) (i 0))
    (while (and  (< i (length a)) (< chra (elt a i)))
      (setq out (+ (* 26 out)  (- (logand lcmask (elt a i)) chra))) ;; -33 is mask to change case
      (setq i (+ 1 i)))
    ))

(defun ss-col-letter (a)
  "returns the letter of the column id arg -- expects int"
  (let ((out "") (n 1) (chra (string-to-char "A")))
    (while (<= 0 a)
      (setq out (concat (char-to-string (+ chra (% a 26))) out))
      (setq a  (- (/ a 26) 1)) )
    out ))

(defun ss-edit-format ( )
  "edit the format of the selected cell" (interactive)
  (let*  (  (current-cell (concat (ss-col-letter ss-cur-col) (int-to-string ss-cur-row)))
            (m (avl-tree-member ss-data (ss-addr-to-index current-cell)))
            (of (if m (elt m ss-c-fmt) ss-default-number-fmt))
            (prompt (concat "Cell " current-cell ": "))
            (nt (read-string prompt of  ss-format-history )))
    (if m
        (progn
          (aset m ss-c-fmla nt)
          (aset m ss-c-fmtd (ss-format-number nt (elt m ss-c-val)))
          (ss-draw-cell ss-cur-col ss-cur-row (elt m ss-c-fmtd)))
      (progn
        (setq m (ss-new-cell current-cell))
        (aset m ss-c-fmla nt)))))

(defun ss-set-mark () "set the mark to the current cell" (interactive)
       (setq ss-mark-cell ss-current-cell))

(defun ss-edit-cell ( )
  "edit the selected cell"
  (interactive)

  (let*  (  (current-cell (concat (ss-col-letter ss-cur-col) (int-to-string ss-cur-row)))
            (prompt (concat "Cell " current-cell ": "))
            (m (avl-tree-member ss-data (ss-addr-to-index current-cell)))
            (ot (if m (if (string= "" (elt m ss-c-fmla)) (elt m ss-c-val) (elt m ss-c-fmla)) "" ))
            (nt 1) )
    (setq nt (if (equal 'return last-input-event)
                 (read-string prompt ot  ss-input-history )
               (read-string prompt (char-to-string last-input-event)  ss-input-history )))

    (ss-update-cell current-cell nt)
    (ss-move-down)
    ))


(defun ss-update-cell (current-cell nt)
  "Update the value/formula  of current cell to nt"
  (let  ( (m (avl-tree-member ss-data (ss-addr-to-index current-cell)))
          )

    ;;delete this cell from its old deps
    (if m (let ((refs (ss-formula-cell-refs  (aref m ss-c-fmla))))
            (dolist (ref refs)
              (ss-del-dep (elt ref 0) current-cell))
            ) nil )
    ;; if this is a formula, eval and do deps
    (if (and (< 0 (length nt)) (= (string-to-char "=") (elt nt 0))) ; formulas start with  =
        (progn
          (if m
              (aset m ss-c-fmla nt)
            (progn
              (setq m (ss-new-cell current-cell))
              (aset m ss-c-fmla nt)
              (avl-tree-enter ss-data m)
              ))

          (let ((s (concat nt " ")) (deps (list)))

            (setq s (replace-regexp-in-string ss-cell-or-range-re (lambda (str)
                                                                    (let ((addr (match-string 1 str)))
                                                                      (push addr deps)
                                                                      (ss-cell-val addr))) s nil nil 1))
            (dolist (dep deps)
              (ss-add-dep dep current-cell))
            (setq nt (calc-eval (substring s 1 -1 ) ))  ;; throw away = and  ' '
            ))
      (if m (aset m ss-c-fmla "") nil))
    ;; nt is now not a formula
    (if m
        (progn
          (aset m ss-c-val nt)
          (aset m ss-c-fmtd (ss-format-number (elt m ss-c-fmt) nt)))
      (progn  ;;else
        (setq m (ss-new-cell current-cell))
        (aset m ss-c-val nt)
        (aset m ss-c-fmtd (ss-format-number ss-default-number-fmt nt))
        (avl-tree-enter ss-data m)
        ))
    ;;do eval-chain
    (dolist (a (aref m ss-c-deps))
      (if (equal a current-cell)
          nil ;; don't loop infinite
        (ss-eval-chain a current-cell)))

    ;;draw it
    (ss-draw-cell ss-cur-col ss-cur-row (ss-highlight (ss-pad-right (elt m ss-c-fmtd) ss-cur-col) ))
    ))


(defun ss-xml-query (node child-node)
  "search an xml doc to find nodes with tags that match child-node"
  (let ((match (list (if (string= (car node) child-node) node nil)))
        (children (cddr node)))
    (if (listp children)
        (dolist (child children)
          (if (listp child)
              (setq match (append match (ss-xml-query child child-node))) nil)
          ) nil )
    (delq nil match)))

(defun ss-load (filename)
  "Load a file into the Spreadsheet."
  (interactive "FFilename:")
  (let ((patterns (list (cons "\\.\\(xlsx\\|XLSX\\)$" 'ss-load-xlsx)
                        (cons "\\.\\(csv\\|CSV\\)$" 'ss-load-csv)
                        )))
    (dolist (e patterns)
      (if (string-match-p (car e) filename)
          (eval (list (cdr e) filename))
        nil) )))

(defun ss-load-xlsx (filename )
  "Try and read an XLSX file into ss-mode"
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
      (dolist (sh (ss-xml-query (car xml) "Override"))
        (let ((type (cdr (assoc 'ContentType (elt sh 1))))
              (name (cdr (assoc 'PartName (elt sh 1)))))
          (if (string= type sheettype) (push name sheets))
          (if (string= type styletype) (push name styles))
          (if (string= type stringstype) (push name strings))
          )))

    (setq sheets (vconcat (nreverse sheets)))
    (setq styles (nreverse styles))
    (setq strings (nreverse strings))
    (dolist (sfn sheets)

          (dolist (stringfn strings)
            (with-temp-buffer
              (erase-buffer)
              (shell-command (concat "unzip -p " filename " " (substring stringfn 1)) (current-buffer))
              (setq xml (xml-parse-region (buffer-end -1) (buffer-end 1)))
              (let ((ts (ss-xml-query (car xml) "t")) )
                (dolist (cell ts)
                  (push (elt cell 2) shstrs)
                  ))))
          (setq shstrs (vconcat  (nreverse shstrs)))
          ;; load styles...
          (with-temp-buffer
            (erase-buffer)
            (shell-command (concat "unzip -p " filename " " (substring sfn 1)) (current-buffer))
            (setq xml (xml-parse-region (buffer-end -1) (buffer-end 1)))
            (let ((sheet (avl-tree-create 'ss-avl-cmp))
		  (cols (ss-xml-query (car xml) "col")))
	      (push ss-sheets sheet)
	      (setq ss-data sheet)
              (dolist (col cols)
                (let ((max (string-to-int (cdr (assoc 'max (elt col 1)))))
                      (min (string-to-int (cdr (assoc 'min (elt col 1)))))
                      (width (cdr (assoc 'width (elt col 1))))
                      (len (length ss-col-widths)) )
                  (if (< len  max)
                      (setq ss-col-widths (vconcat ss-col-widths (make-vector (- max len) (truncate width)))) nil )
                  (dotimes (i (+1 (- max min)))
                    (aset ss-col-widths (+ min i) (truncate width)))
                  )))
            (dolist (cell (ss-xml-query (car xml) "c"))
              (let* ((range (prin1-to-string (cdr (assoc 'r (elt cell 1))) t))
                     (value (prin1-to-string (car (cddr (assoc 'v cell))) t))
                     (fmla (prin1-to-string  (car (cddr (assoc 'f cell))) t)))

                (if (string= value "nil") (setq value ""))
                (if (string= "s" (cdr (assoc 't (elt cell 1)))) (setq value (elt shstrs (string-to-int value))))
                (if (string= fmla "nil") (setq fmla ""))

                (if (string-match "\\([A-Za-z]+\\)\\([0-9+]\\)" range)
                    (let ((row (string-to-int (match-string 2 range)))
                          (col (ss-col-number (match-string 1 range))))
                      (if (<= ss-max-row row) (setq ss-max-row (+ 2 row)))
                      (if (<= ss-max-col col)
                          (progn
                            (setq ss-col-widths (vconcat ss-col-widths (make-vector  (- col ss-max-col ) 7)))
                            (setq ss-max-col (length ss-col-widths))
                            ) nil )
                      ) nil)
                (if (string= "" fmla)
                    (ss-update-cell range value)
                  (ss-update-cell range (concat "=" fmla)))
                ))))
    (setq ss-data (car sheets))
    (ss-draw-all)
      ))

(defun ss-load-csv (filename)
  "Read a CSV file into ss-mode"
  (interactive "fFilename:")
  (let ((mybuff (current-buffer))
        (cc ss-cur-col)
        (cr ss-cur-row))
    
    (with-temp-buffer
     (insert-file-contents filename)
     (beginning-of-buffer)
     (let ((x 0) (cell "") (cl ss-cur-col)
           (rw ss-cur-row) (j 0)
           (re ",?\\(\"\\(\\(\"\"\\|[|^\"]\\)+\\)\"\\|\\([^,]+\\)\\),?")
           (newline ""))
       (while (not (eobp))
         (setq newline (thing-at-point 'line))
         (setq j 0)
         (while (and (< j (length newline)) (string-match re newline j))
           (setq x (if (match-string 2 newline) 2 1))
           (setq cell (match-string x newline))
           (setq j (+ 1 (match-end x)))
           (with-current-buffer mybuff
             (ss-update-cell (concat (ss-col-letter cl) (int-to-string rw)) cell)
             )
           (setq cl (+ cl 1))
           (if (<= ss-max-col cl)
               (progn
                 (setq ss-max-col (+ cl 1))
                 (setq ss-col-widths (vconcat ss-col-widths (list (elt ss-col-widths (- cl 1)))))
                 ) nil ))
         (setq rw (+ rw 1))
         (if (<=  ss-max-row  rw)
             (setq ss-max-row (+ rw 1)) nil)
         (setq cl ss-cur-col)
         (forward-line 1)
         )))
    (setq ss-cur-col cc)
    (setq ss-cur-rw cr)
    (ss-draw-all)
    ))


;;  ____                     _
;; |  _ \ _ __ __ ___      _(_)_ __   __ _
;; | | | | '__/ _` \ \ /\ / / | '_ \ / _` |
;; | |_| | | | (_| |\ V  V /| | | | | (_| |
;; |____/|_|  \__,_| \_/\_/ |_|_| |_|\__, |
;;                                   |___/
;; functions dealing with the cursor and cell drawing /padding



(defun ss-format-number (fmt val)
  "output text format of numerical value according to format string

000000.00  -- pad to six digits (or however many zeros to left of .) round at two digits, or pad out to two
#,###.0 -- insert a comma (anywhere to the left of the . is fine -- you only need one and it'll comma ever three
0.00## -- Round to four digits, and pad out to at least two."
  (if (stringp val) nil (setq val (number-to-string val)))
  (if (equal 0 (string-match "^ *\\+?-?[0-9,\\.]+ *$" val))
      (progn
        (let ((ip 0)   ;;int padding
              (dp 0)   ;;decmel padding
              (dr 0)   ;;decmil round
              (dot-index (string-match "\\." fmt))
              (power 0)
              (value (if (numberp val) val (string-to-number val)))
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
                    (rep (lambda (a) (concat (match-string 1 a) (match-string 4 a) "," (match-string 5 a) (match-string 2 a) (match-string 3 a)))))
                (while (string-match re dec)
                  (setq dec (replace-regexp-in-string re rep dec)))
                ) nil )
          dec))
    val))

(defun ss-highlight (txt)
  "highlight text"

  (let ((myface '((:foreground "White") (:background "Blue"))))
    (propertize txt 'font-lock-face myface)))
;;  (concat ">" txt "<"))

(defun ss-pad-center (s i)
  "pad a string out to center it - expects stirng, col no (int)"
  (setq s (replace-regexp-in-string "\n" " " s))
  (let ( (ll (length s) ))
    (if (>= ll (- (elt ss-col-widths i) 2))
        (make-string (elt ss-col-widths i) (string-to-char "#"))
      (let* ( (pl (/ (- (elt ss-col-widths i) ll)  2))  ; half the pad length
              (pad (make-string (- (elt ss-col-widths i) pl 1 ) (string-to-char " "))) )
        (concat (make-string pl (string-to-char " ")) s pad)) )
    ))

(defun ss-pad-left (s i)
  "pad to the left  - expects stirng, col no (int)"
  (setq s (replace-regexp-in-string "\n" " " s))
  (let ( (ll (length s)) )
    (if (>= ll (- (elt ss-col-widths i) 2))
        (make-string (elt ss-col-widths i) (string-to-char "#"))
      (concat s (make-string (- (elt ss-col-widths i)  ll) (string-to-char " ")))
      )))

(defun ss-pad-right (s i)
  "pad to the right  - expects stirng, col no (int)"
  (setq s (replace-regexp-in-string "\n" " " s))
  (let ( (ll (length s)) )
    (if (>= ll (- (elt ss-col-widths i) 2))
        (make-string (elt ss-col-widths i) (string-to-char "#"))
      (concat (make-string (- (elt ss-col-widths i)  ll) (string-to-char " ")) s)
      )))


(defun ss-draw-all ()
  "Populate the current ss buffer." (interactive)
  ;;  (pop-to-buffer ss-empty-name nil)

  (let ((i 0) (j 0) (k 0) (header (make-string ss-row-padding (string-to-char " "))))
    (beginning-of-buffer)
    (erase-buffer)
    (dotimes (i (- ss-max-col 1))  ;; make header
      (setq header (concat header (ss-pad-right (ss-col-letter i) i))))
    (dotimes (j ss-max-row) ;; draw buffer
      (if (= 0 j)
          (insert  (ss-highlight header))
        (progn
          (insert (ss-highlight (format (concat "%" (int-to-string ss-row-padding) "d") j) ))
          (dotimes (i ss-max-col)
            (let ((m (avl-tree-member ss-data (+ j (* ss-max-row (+ i 1))))))
              (if m
                  (insert (ss-pad-right (elt m ss-c-fmtd) i))
                (insert (make-string (elt ss-col-widths i) (string-to-char " "))))))))

      (insert "\n"))
    (set-buffer-modified-p nil)
    (ss-move-cur-cell 0 0) ;;draw cursor
    ))


(defun ss-draw-cell (x y text)
  "redraw one cell on the ss  - expects  col no. (int), row no, (int), cell value (padded string)"
  (let ((col ss-row-padding) (i 0))
    ;;    (pop-to-buffer ss-empty-name nil)
    (setq cursor-type nil)  ;; no cursor
    (setq truncate-lines 1)  ;; no wrap-around

    (dotimes (i x) (setq col (+ col (elt ss-col-widths i))) )
    (goto-line (+ y 1))
    (move-to-column col)

    (insert text)
    (delete-forward-char (length text))
    (recenter)
    ))



(defun ss-move-left ()
  (interactive)
  (if (> ss-cur-col  0)
      (ss-move-cur-cell -1 0) nil)
  )

(defun ss-move-right ()
  (interactive)
;  (debug)
  (if (<= ss-cur-col (- ss-max-col 3))
      (ss-move-cur-cell 1 0)
    (progn
      (setq ss-max-col (+ 1 ss-max-col))
      (setq ss-col-widths   (vconcat ss-col-widths (list (elt ss-col-widths ss-cur-col))))
      (setq ss-cur-col (+ 1 ss-cur-col))
      (ss-draw-all)
      (ss-move-cur-cell 1 0)
      )))

(defun ss-move-up ()
  (interactive)
  (if (> ss-cur-row 1)
      (ss-move-cur-cell 0 -1)
    nil ))

(defun ss-move-down ()
  (interactive)
  (if (< ss-cur-row (- ss-max-row 1))
      (ss-move-cur-cell 0 1 )
    (progn
      (setq ss-max-row (+ 1 ss-max-row))
      (setq ss-cur-row (+ 1 ss-cur-row))
      (ss-draw-all)
      (ss-move-cur-cell 1 0)
      )))

(defun ss-move-cur-cell (x y) (interactive)
  "Move the cursor or selection overaly"

  (let* ((new-row (+ ss-cur-row y))
	 (new-col (+ ss-cur-col x))
	 (n (avl-tree-member ss-data (+ new-row (* ss-max-row (+ 1 new-col)))))
	 (row-pos (line-beginning-position (- (line-number-at-pos) new-row)))
	 (col-pos (apply '+ (mapcar (lambda (x) (elt ss-col-widths x)) (number-sequence 0 new-col))))) ;; add col-widts b/t 0 and col
    (if (overlayp ss-cursor)
	(move-overlay ss-cursor (+ row-pos col-pos) (+ row-pos col-pos (elt ss-col-widths (+ 1 new-col))))
      (progn
;;	(debug)
	(setq ss-cursor (make-overlay (+ row-pos col-pos) (+ row-pos col-pos (elt ss-col-widths (+ 1 new-col)))))
	(overlay-put ss-cursor 'face '((:foreground "White") (:background "Blue")))
	))
    (setq ss-cur-row new-row)
    (setq ss-cur-col new-col)
    (minibuffer-message (concat "Cell " (ss-col-letter ss-cur-col) (int-to-string ss-cur-row)
                                       " : " (if n (if (string= "" (elt n ss-c-fmla)) (elt n ss-c-val ) (elt n ss-c-fmla)) "" )))

    ))


(defun ss-increase-cur-col ()
  "Increase the width of the current column" (interactive)
  (aset ss-col-widths ss-cur-col (+ 1 (elt ss-col-widths ss-cur-col)))
  (ss-draw-all))

(defun ss-decrease-cur-col ()
  "Decrease the width of the current column" (interactive)
  (aset ss-col-widths ss-cur-col (- (elt ss-col-widths ss-cur-col) 1))
  (ss-draw-all))

;;   ____     _ _   _____            _             _   _
;;  / ___|___| | | | ____|_   ____ _| |_   _  __ _| |_(_) ___  _ __
;; | |   / _ \ | | |  _| \ \ / / _` | | | | |/ _` | __| |/ _ \| '_ \
;; | |__|  __/ | | | |___ \ V / (_| | | |_| | (_| | |_| | (_) | | | |
;;  \____\___|_|_| |_____| \_/ \__,_|_|\__,_|\__,_|\__|_|\___/|_| |_|
;; fuctions dealing with eval


(defun ss-cell-val (address)
  "get the value of a cell"
  (let ((addr (replace-regexp-in-string "\\$" "" address)))
    (if (string-match ss-range-parts-re addr)
        (progn
          (let* ((s (ss-addr-to-index (concat  (match-string 1 addr) (match-string 2 addr))))
                 (cl (- (ss-addr-to-index (concat  (match-string 1 addr) (match-string 4 addr))) s))
                 (rl (- (ss-addr-to-index (concat  (match-string 3 addr) (match-string 2 addr))) s))
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
            (setq rl (/ rl ss-max-row))
            (dotimes (r (+ 1 rl))
              (setq rv (concat rv "["))
              (dotimes (c (+ 1 cl))
                (setq m (avl-tree-member ss-data (+ s c (* ss-max-row r))))
                (setq rv (concat rv (if m (elt m ss-c-val) "0") (if (= c cl) "]" ",") ))
                )
              (setq rv (concat rv (if (= r rl) "]" ",")))
              ) rv))
      (progn
        (let ((m (avl-tree-member ss-data  (ss-addr-to-index addr))) )
          (if m (elt m ss-c-val) "0")
          )))))


(defun ss-eval-chain (addr chain)
  "Update cell and all deps"
  (let ((m (avl-tree-member ss-data (ss-addr-to-index addr))) )
    (if m
        (progn
          (ss-eval-fun addr)
          (dolist (a (elt m ss-c-deps))
            (if (member a chain) nil
              (ss-eval-chain a (append chain (list addr)))
              ))))))


(defun ss-formula-cell-refs (formula)
  "Takes a string and finds all the cell references in a string Ex:
   (ss-fomual-cell-refs (\"= A1+B1\"))
   returns a list ((\"A1\" 2 3) (\"B1\" 5 6))"
  (let* ((s (concat formula " "))
         (j 0) (retval (list )))
    (while (string-match ss-cell-or-range-re s j)
      (setq retval (append retval (list (list (match-string 1 s) (match-beginning 1) (match-end 1)))))
      (setq j (+ 1(match-end 1)))
      )
    retval))

(defun ss-eval-fun (addr)
  "sets cell value based on its function; draws"
  (let*  ( (m (avl-tree-member ss-data (ss-addr-to-index addr)))
           (cv "")
           (s (if m (elt m ss-c-fmla) ""))
           (refs (ss-formula-cell-refs s)) )
    (if refs
        (progn
          (dolist (ref refs)
            (setq cv (ss-cell-val (elt ref 0)))
            (setq s (concat (substring s 0 (elt ref  1)) cv (substring s (elt ref 2)))))
          (setq cv (calc-eval (substring s 1)))
          (aset m ss-c-val cv)
          (aset m ss-c-fmtd (ss-format-number (elt m ss-c-fmt) cv))
          (let ((c 0) (i 0) (chra (- (string-to-char "A") 1)))
            (while (and  (< i (length addr)) (< chra (elt addr i)))
              (setq c (+ (* 26 c)  (- (logand -33 (elt addr i)) chra))) ;; -33 is mask to change case
              (setq i (+ 1 i)))
            (ss-draw-cell (- c 1) (string-to-int (substring addr i))
                          (ss-pad-right (elt m ss-c-fmtd) (- c 1))) )
          cv ) (aref m ss-c-val)
          )))


(defun ss-add-dep (ca cc)
  "Add to dep list.
   CA -- addr to to add
   CC -- cell who depends
"
  (let ((addr (replace-regexp-in-string "\\$" "" ca)))
    (if (string-match ss-range-parts-re addr)
        (progn
          (let* ((s (ss-addr-to-index (concat  (match-string 1 addr) (match-string 2 addr))))
                 (cl (- (ss-addr-to-index (concat  (match-string 1 addr) (match-string 4 addr))) s))
                 (rl (- (ss-addr-to-index (concat  (match-string 3 addr) (match-string 2 addr))) s))
                 )
            (if (< cl 0)
                (progn
                  (setq cl (* -1 cl))
                  (setq s (- s cl)) ) nil)
            (if (< rl 0)
                (progn
                  (setq rl (* -1 rl))
                  (setq s (- s rl))) nil )
            (setq rl (/ rl ss-max-row))
            (dotimes (r (+ 1 rl))
              (dotimes (c (+ 1 cl))
                (ss-add-dep  (ss-index-to-addr (+ s c (* r ss-max-row))) cc)
                )
              )))
      (let  ( (m (avl-tree-member ss-data (ss-addr-to-index addr))))
        (if m
            (aset m ss-c-deps (append (elt m ss-c-deps) (list cc)))
          (progn
            (setq m (ss-new-cell addr))
            (aset m ss-c-deps (list cc))
            (avl-tree-enter ss-data m)
            ))))))

(defun ss-del-dep (ca cc)
  "Remove from dep list."
  (let ((addr (replace-regexp-in-string "\\$" "" ca)))
    (if (string-match ss-range-parts-re addr)
        (progn
          (let* ((s (ss-addr-to-index (concat  (match-string 1 addr) (match-string 2 addr))))
                 (cl (- (ss-addr-to-index (concat  (match-string 1 addr) (match-string 4 addr))) s))
                 (rl (- (ss-addr-to-index (concat  (match-string 3 addr) (match-string 2 addr))) s))
                 )
            (if (< cl 0)
                (progn
                  (setq cl (* -1 cl))
                  (setq s (- s cl)) ) nil)
            (if (< rl 0)
                (progn
                  (setq rl (* -1 rl))
                  (setq s (- s rl))) nil )
            (setq rl (/ rl ss-max-row))
            (dotimes (r (+ 1 rl))
              (dotimes (c (+ 1 cl))
                (ss-del-dep (ss-index-to-addr (+ s c (* r ss-max-row))) cc)
                )
              )))
      (let  ( (m (avl-tree-member ss-data (ss-addr-to-index addr))))
        (if m
            (delete cc (aref m ss-c-deps))
          nil) ))))



;;;###autoload
(define-derived-mode ss-mode text-mode ss-empty-name
  "ss game mode
  Keybindings:
  \\{ss-map} "
  (use-local-map ss-map)

  ;; (unless (featurep 'emacs)
  ;;   (setq mode-popup-menu
  ;;         '("ss Commands"
  ;;           ["Start new game"        ss-start-game]
  ;;           ["End game"                ss-end-game
  ;;            (ss-active-p)]
  ;;           ))



  (setq ss-cur-col 0)
  (setq ss-max-col 3)
  (setq ss-col-widths (make-vector ss-max-col 7))
  (setq ss-cur-row 1)
  (setq ss-max-row 3)
  (setq ss-row-padding 4)
  (setq ss-data (avl-tree-create 'ss-avl-cmp))
  (if (file-exists-p (buffer-file-name)) (ss-load (buffer-file-name)) nil)

  (ss-draw-all)

  )

;;  ____        __ __  __       _   _
;; |  _ \  ___ / _|  \/  | __ _| |_| |__  ___
;; | | | |/ _ \ |_| |\/| |/ _` | __| '_ \/ __|
;; | |_| |  __/  _| |  | | (_| | |_| | | \__ \
;; |____/ \___|_| |_|  |_|\__,_|\__|_| |_|___/

(defmath SUM ( x)
  "add the items in the range"
  (interactive 1 "sum")
  :" vflat(x) * ((vflat(x) * 0 ) + 1)"

  )


(provide 'ss-mode)
(push (cons "\\.\\(xlsx\\|XLSX\\)$" 'ss-mode) auto-mode-alist)
(push (cons "\\.\\(csv\\|CSV\\)$" 'ss-mode) auto-mode-alist)

;;; ss-mode.el enxds here

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
(defvar qs-format-re "\\([-b]\\)\\([#,]*0*.?0*#*\\) ?\\([\\%kMBE]\\)? *")

;;       CELL Format -- [ "A1" 0.5 "= 1/2" "#0.00M" "= 3 /2" (list of cells to calc when changes)]
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
(defvar qs-default-number-fmt "-#,##0.########")
(defvar qs-cursor nil)
(defvar qs-copy-buffer ())

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
(define-key qs-map [deletechar] 'qs-clear-area)
(define-key qs-map [backspace]  'qs-clear-area)
(define-key qs-map [remap set-mark-command] 'qs-set-mark)
(define-key qs-map [remap self-insert-command] 'qs-edit-cell)
(define-key qs-map (kbd "C-x f") 'qs-edit-format)
(define-key qs-map (kbd "C-x r") 'qs-resize-col-to-fit)
(define-key qs-map (kbd "C-x d") 'qs-redraw)
(define-key qs-map (kbd "C-w") 'qs-cut-area)
(define-key qs-map (kbd "M-w") 'qs-copy-area)
(define-key qs-map (kbd "C-y") 'qs-paste-area)


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



(defun qs-clear-area () "delete current cell or marked range" (interactive)
       (let* ((start-col (if qs-mark-cell (if (< qs-cur-col (elt qs-mark-cell 0)) qs-cur-col (elt qs-mark-cell 0)) qs-cur-col))
              (col-len (if qs-mark-cell (- (+ 1 qs-cur-col (car qs-mark-cell)) (* 2 start-col)) 1))
              (start-row (if qs-mark-cell (if (< qs-cur-row (elt qs-mark-cell 1)) qs-cur-row (elt qs-mark-cell 1)) qs-cur-row))
              (row-len (if qs-mark-cell (- (+ 1 qs-cur-row (elt qs-mark-cell 1)) (* 2 start-row)) 1))
              (rows ()))
         (dotimes (y row-len)
           (dotimes (x col-len)
             (let* ((del-cell-addr (qs-rowcol-to-addr (+ x start-col) (+ y start-row)))
                    (del-cell-index (qs-rowcol-to-index (+ x start-col) (+ y start-row)))
                    )
               (qs-update-cell del-cell-addr "" "")
               )))
         (setq qs-mark-cell nil)
         (qs-draw-all)
         (set-buffer-modified-p 1)
         )
       )

(defun qs-cut-area () "cut currrent cell or marked area" (iteractive)
       (qs-copy-area)
       (qs-clear-area)
       (setq qs-mark-cell nil)
       (qs-draw-all)
       )

(defun qs-copy-area () "copy current cell or marked range into the copy buffer" (interactive)
       (let* ((start-col (if qs-mark-cell (if (< qs-cur-col (elt qs-mark-cell 0)) qs-cur-col (elt qs-mark-cell 0)) qs-cur-col))
              (col-len (if qs-mark-cell (- (+ 1 qs-cur-col (car qs-mark-cell)) (* 2 start-col)) 1))
              (start-row (if qs-mark-cell (if (< qs-cur-row (elt qs-mark-cell 1)) qs-cur-row (elt qs-mark-cell 1)) qs-cur-row))
              (row-len (if qs-mark-cell (- (+ 1 qs-cur-row (elt qs-mark-cell 1)) (* 2 start-row)) 1))
              (this-cell "")
              (rows ()))
         (dotimes (y row-len)
           (let ((ccols ()))
             (dotimes (x col-len)
               (setq this-cell (qs-rowcol-to-index (+ x start-col) (+ y start-row)))
               (setq ccols (append ccols (list (avl-tree-member qs-data this-cell))))
               )
             (setq rows (append rows (list ccols)))
             )
           )
         (setq qs-copy-buffer rows)
         (setq qs-mark-cell nil)
         ))

(defun qs-transform-fmla (from-addr to-addr fmla)
  "transform a function from from to to moving all addresses relative to the addresses
EX:  From: A1 To: B1 Fun: = A2 / B1
     Returns: = B2 / C1 "
  (let* ((from-rc (qs-addr-to-rowcol from-addr))
         (to-rc (qs-addr-to-rowcol to-addr))
         (col-delta (- (elt to-rc 0) (elt from-rc 0)))
         (row-delta (- (elt to-rc 1) (elt from-rc 1)))
         (cell-ref-re "\\($?[A-Za-z]+\\)\\($?[0-9]+\\)")
         (i 0)
         (flen (length fmla))
         (new-fmla "")
         (end-flma "")
         (worked t)
         )

    (while (and (< i (length fmla)) (string-match cell-ref-re fmla i))
      (let* ((old-addr-col (match-string 1 fmla))
             (old-addr-row (match-string 2 fmla))
             (old-col-row (qs-addr-to-rowcol (string-replace "$" "" (concat old-addr-col old-addr-row))))
             (newcol (if (not (equal (string-to-char "$") (elt old-addr-col 0))) (+ (elt old-col-row 0) col-delta) (elt old-col-row 0)))
             (newrow (if (not (equal (string-to-char "$") (elt old-addr-row 0))) (+ (elt old-col-row 1) row-delta) (elt old-col-row 1)))
             (mat-end (match-end 0 ))
             )
        (if (or (not worked) (< newcol 0) (< newrow 1))
            (progn
              (setq worked nil)
              (setq i (+ 1 (length fmla))) )
          (progn
            (setq new-fmla (concat new-fmla
                                   (substring fmla i (match-beginning 0 ))
                                   (if (equal (string-to-char "$") (elt old-addr-col 0)) "$")
                                   (qs-col-letter newcol)
                                   (if (equal (string-to-char "$") (elt old-addr-row 0)) "$")
                                   (int-to-string newrow)))
            (if (and mat-end (< mat-end flen)) (setq end-fmla (substring fmla mat-end flen)))
            (setq i mat-end)
            ) )
        ))
    (if end-fmla  (setq new-fmla (concat new-fmla end-fmla)))
    (if worked new-fmla worked)
    ))



(defun qs-paste-area () "paste the copy buffer into the sheet" (interactive)
       (if qs-copy-buffer
           (let* (
                  (paste-rows (length qs-copy-buffer))
                  (paste-cols (length (elt qs-copy-buffer 0)))
                  (start-col (if qs-mark-cell (if (< qs-cur-col (elt qs-mark-cell 0)) qs-cur-col (elt qs-mark-cell 0)) qs-cur-col))
                  (col-len (if qs-mark-cell (- (+ 1 qs-cur-col (car qs-mark-cell)) (* 2 start-col)) 1))
                  (start-row (if qs-mark-cell (if (< qs-cur-row (elt qs-mark-cell 1)) qs-cur-row (elt qs-mark-cell 1)) qs-cur-row))
                  (row-len (if qs-mark-cell (- (+ 1 qs-cur-row (elt qs-mark-cell 1)) (* 2 start-row)) 1))
                  (copy-elt-y 0)
                  (copy-elt-x 0)
                  )
             (if (equal col-len 1) (setq col-len paste-cols))
             (if (equal row-len 1) (setq row-len paste-rows))
             (dotimes (y row-len)
               (let ((row (elt qs-copy-buffer copy-elt-y)) )
                 (dotimes (x col-len)
                   (let ((cell (elt row copy-elt-x)) )
                     (if cell
                         (let* (
                                (addr (qs-rowcol-to-addr  (+ x start-col) (+ y start-row)))
                                (cell-fmla  (elt cell qs-c-fmla))
                                (fmla (if (or (not cell-fmla) (string= "" cell-fmla)) ""
                                        (qs-transform-fmla (elt cell qs-c-addr) addr cell-fmla)))
                                (txt  (if (and fmla (not (string= "" fmla))) fmla (elt cell qs-c-val)))
                                (fmt  (if (string= "" (elt cell qs-c-fmt)) nil (elt cell qs-c-fmt)))
                                (m (qs-update-cell addr txt fmt))
                                )
                           (setq copy-elt-x (if (eql  (+ 1 copy-elt-x) paste-cols) 0 (+ 1 copy-elt-x)))
                           ))))
                 (setq copy-elt-y (if (eql (+ 1 copy-elt-y)  paste-rows ) 0 (+ 1 copy-elt-y)))
                 ))
             )
         )
       (setq qs-mark-cell nil)
       (qs-draw-all)
       )


(defun qs-avl-cmp (a b)
  "This is the function used by avl tree to compare ss addresses"
  (let ((A (if (sequencep a) (elt a qs-c-addr) a)) (B (if (sequencep b) (elt b qs-c-addr) b)))  ; a or b can be vectors or addresses
    (< (qs-addr-to-index A) (qs-addr-to-index B)) ))


(defun qs-rowcol-to-addr ( col row ) "return the cell address (eg A1) form row and col)"
       (concat (qs-col-letter  col) (int-to-string row)))

(defun qs-rowcol-to-index (col row) "return the index for row col"
       (+ (* qs-max-row (+ 1 col)) row))

(defun qs-addr-to-rowcol (a) "return (col row) from addr"
       (let ((chra (string-to-char "A") )
             (col 0)
             (row 0)
             (i 0))
         (if (sequencep a)
             (progn
               (while (and  (< i (length a)) (<= chra (elt a i)))
                 (setq col (+ (* 26 col)  (- (logand qs-lcmask (elt a i)) chra))) ;; -33 is mask to change case
                 (setq i (+ 1 i)))
               (setq row (floor (string-to-number (substring a i))))
               (list col row))
           nil )))


(defun qs-addr-to-index (a)  "Convert from ss addr (e.g. A1) to index  -- expects [A-Z]+[0-9]+"
       (let ((chra (- (string-to-char "A") 1)))
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
            (prompt (concat "Cell " current-cell ": " ))
            (newfmt (read-string prompt of qs-format-history )))
                                        ;    (add-to-history qs-format-history newfmt)
    (if (string-match qs-format-re newfmt)
        (if m
            (progn
              (aset m qs-c-fmt newfmt)
              (aset m qs-c-fmtd (qs-format-number newfmt (elt m qs-c-val)))
              (qs-draw-all))
          (progn
            (setq m (qs-new-cell current-cell))
            (aset m qs-c-fmt newfmt)
            ))
      (message "Invalid format!!")
      )))

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
               (read-string prompt (char-to-string last-input-event) qs-input-history )))
                                        ;  (add-to-history qs-input-history nt)
    (setq m (qs-update-cell current-cell nt))
    (qs-draw-cell qs-cur-col qs-cur-row (qs-pad-right (elt m qs-c-fmtd) qs-cur-col))
    (qs-move-down)
    ))
(defun qs-default-fmt (num ) "get a default format for a number"
       (let ((nnt (if (numberp num) num (qs-s-to-n num))))
         (cond
          ((> .001 nnt) "-0.00E")
          ((> 1 nnt) "-0.00%")
          ((< (expt 10 12) nnt) "-0.00E")
          ((< (expt 10 9) nnt) "b#,##0.#B")
          ((< (expt 10 6) nnt) "b#,##0.#M")
          ((< 1000 nnt) "b#,##0.#k")
          (t qs-default-number-fmt)
          ))
       )

(defun qs-update-cell (current-cell nt &optional format)
  "Update the value/formula  of current cell to nt"
  (let  ( (m (avl-tree-member qs-data (qs-addr-to-index current-cell)))
          )
    ;;delete this cell from its old deps
    (if m (let ((refs (qs-formula-cell-refs  (aref m qs-c-fmla))))
            (dolist (ref refs)
              (qs-del-dep (elt ref 0) current-cell))
            ))
    ;; if this is a formula, eval and do deps
    (if (and (< 0 (length nt)) (string= "=" (substring nt 0 1)))
        (progn  ;;If there is a formula set formula and calc it.
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

            (dolist (dep deps)  ;; set in the above lambda
              (qs-add-dep current-cell dep))

            (setq nt (qs-calc-eval-error (calc-eval (substring s 1 (length s)))))

            (if format
                (aset m qs-c-fmt format)
              (let ( (nnt (qs-s-to-n nt)))
                (if nnt
                    (aset m qs-c-fmt (qs-default-fmt nnt))
                  ))
              )
            )
          )
      (if m (aset m qs-c-fmla "") nil)
      )
    ;; set the cval and c-fmtd (for formulas, nt was reset to c-val in fmla above)
    (if m
        (let ( (nnt (qs-s-to-n nt)) )
          (aset m qs-c-val (if nnt (number-to-string nnt) nt))
          (aset m qs-c-fmtd (qs-format-number (elt m qs-c-fmt) (if nnt nnt nt)))


          ;;do eval-chain
          (let ((cell-deps (aref m qs-c-deps)))
            (dolist (dep cell-deps)
              (if (not (string= (upcase dep) (upcase current-cell)))
                  (qs-eval-chain dep (list (upcase current-cell)))
                )
              )
            )


          )
      (progn
        (setq m (qs-new-cell current-cell))
        (let ( (nnt (qs-s-to-n nt)))
          (if nnt
              (let ((nf (qs-default-fmt nnt)))
                (if format (setq nf format))
                (aset m qs-c-val (number-to-string nnt))
                (aset m qs-c-fmt nf)
                (aset m qs-c-fmtd (qs-format-number nf nnt ))
                )
            (progn
              (aset m qs-c-fmtd nt)
              (aset m qs-c-val nt)
              (if format
                  (aset m qs-c-fmt format))
              )
            )
          )
        (avl-tree-enter qs-data m)
        )
      )

    m
    ))



(defun qs-xml-query (node child-node)
  "search an xml doc to find nodes with tags that match child-node"
  (let ((match (list (if (string= (car node) child-node) node nil)))
        (children (cddr node)))
    (dolist (child children)
      (if (listp child)
          (setq match (append match (qs-xml-query child child-node))) nil)
      ) nil )
  (delq nil match))

(defun qs-load (filename)
  "Load a file into the Spreadsheet."
  (interactive "FFilename:")
  (let ((patterns (list
                   (cons "\\.XLSX$" 'qs-load-xlsx)
                   (cons "\\.CSV$" 'qs-load-csv)
                   ))
        (done nil))
    (dolist (e patterns)
      (if (not done)
          (setq done  (if (string-match-p (car e) (upper filename))
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
       (setq case-fold-search 1)  ;; don't worry about case
       (if (string-match "\\.csv$" (buffer-file-name))
           (let ((x 0) (cell "") (cl qs-cur-col)
                 (rw qs-cur-row) (j 0)
                 (re " *\\(\"\\([^\"]+\\)\\|\\([^,]+\\)\\) *\"? *[,\n]")
                 (newline "")
                 (split ())
                 (lines ()) )
             (read-only-mode 0)
             (beginning-of-buffer)
             (while (not (eobp))
               (setq newline (thing-at-point 'line))
               (setq split ())
               (setq j 0)
               (if (equal "\n" newline)
                   (add-to-ordered-list 'lines (list nil ) (+ 1 (length lines)))
                 (progn
                   (while (and (< j (length newline)) (string-match re newline j))
		     (setf x 1)
		     (while (match-string (+ 1 x) newline)
		       (setq x (+ 1 x))
		       )
                     (setq cell (string-replace "`" "\"" (match-string x newline)))
                     (setq j (+ 1 (match-end x)))
                     (add-to-ordered-list 'split cell (+ 1 (length split)))
                     )
                   (add-to-ordered-list 'lines split (+ 1 (length lines)))
                   ))
               (forward-line 1)
               )
             (erase-buffer)
             (setq j 1)
             (dolist (line lines)
               (progn
                 (setq x 0)
                 (dolist (cell line)
                   (if cell (progn
                              (setq qs-cur-col x)
                              (setq qs-cur-row j)
                              (qs-update-cell (concat (qs-col-letter x) (int-to-string j)) cell)
                              (setq x (+ x 1))
                              (if (= qs-max-col x)
                                  (progn
                                    (setq qs-max-col (+ x 1))
                                    (setq qs-col-widths (vconcat qs-col-widths (list 5)))
                                    )
                                ))))
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
             )) nil )

(defun qs-draw-csv ()
  "Save to a CSV file"
  (let ((fn (buffer-file-name))
        (empty-lines ()))
    (read-only-mode 0)
    (erase-buffer)
    (dotimes (j qs-max-row) ;; draw buffer
      (let ((line "")
            (empty-cells ()))
        (dotimes (i qs-max-col)
          (let* (
                 (m        (avl-tree-member qs-data (qs-addr-to-index (concat (qs-col-letter i) (int-to-string (+ 1 j))))))
                 (cell-val (if m (if (string= "" (elt m qs-c-fmla)) (elt m qs-c-val) (elt m qs-c-fmla)) nil ))
                 )
            ( if (and m (not (string= "" cell-val)))
                (progn
                  (if (not (string= "" line)) (push line empty-cells))
                  (setq line (string-join (append  empty-cells (list (concat "\"" (string-replace "\"" "`" cell-val) "\"" ))) ","))
                  (setq empty-cells ())
                  )
              (push "" empty-cells)
              )     ))

        (if (not (string= "" line))
            (progn
              (insert (concat (string-join (append empty-lines (list line)) "\n") "\n" ))
              (setq empty-lines ())
              )
          (push "" empty-lines)
          )
        )))
  (read-only-mode 1))
;;  ____                     _
;; |  _ \ _ __ __ ___      _(_)_ __   __ _
;; | | | | '__/ _` \ \ /\ / / | '_ \ / _` |
;; | |_| | | | (_| |\ V  V /| | | | | (_| |
;; |____/|_|  \__,_| \_/\_/ |_|_| |_|\__, |
;;                                   |___/
;; functions dealing with the cursor and cell drawing /padding



(defun qs-format-number (format val)
  "output text format of numerical value according to format string
   000000.00  -- pad to six digits (or however many zeros to left of .) round at two digits, or pad out to two
   #,###.0 -- insert a comma (anywhere to the left of the . is fine -- you only need one and it'll comma every three
   0.00## -- Round to four digits, and pad out to at least two."
  (let ((string-val (if (listp val) " " (if (numberp val) (format "%f1000" val) (if (stringp val) val " ")))))
    (if (string-match "^ *\\+?-?[0-9,\\.]+ *$" string-val)
        (progn
          (string-match qs-format-re format)
          (let (
		(neg-type (match-string 1 format))
		(fmt (match-string 2 format))
                (exp (match-string 3 format))
                (dec "")) ;;formatted string
            (let ((value (* (cond
                             ((equal exp "%") (expt 10 2))
                             ((equal exp "k") (expt 10.0 -3))
                             ((equal exp "M") (expt 10.0 -6))
                             ((equal exp "B") (expt 10.0 -9))
                             (t 1))
                            (float (if (numberp val) val (string-to-number val))))))

              (if (or (equal "E" exp) (equal "e" exp))  ;; scientific notation (1 sig digit)
                  (let ((c 0)
                        (p 1)
                        (v value))
                    (if (= 0 (or v (floor v)))
                        (while (= 0 (or v (floor v)))
                          (progn
                            (setq c (- c 1))
                            (setq p (* 10 p))
                            (setq v (* p value))
                            ))
                      (while (> (floor (* (float value) p )) 10)
                        (progn
                          (setq p (/ (float p) 10.0))
                          (setq c (+ 1 c))
                          )))
                    (setq exp (concat "e"  (if (< 0 c) "+") (int-to-string c)))
                    (setq value (* value (float p)))
                    ))
              (setq dec (number-to-string value))

              (let* ((re "[#,]*\\(0*\\).?\\(0*\\)\\(#*\\)")  ;; Pad out the zeros and round of the #'s
                     (match (string-match re fmt))
                     (lz (length (match-string 1 fmt)))
                     (rz (length (match-string 2 fmt)))
                     (rh (length (match-string 3 fmt)))
                     (se (string-match "\\([0-9]*\\)\\.\\([0-9]*\\)" dec))
                     (ints (match-string 1 dec))
                     (decs (match-string 2 dec))
                     )
                (if (> (length decs) (+ rz rh))  ;;round if decmil too long
                    (let ((overflow (substring decs (+ rz rh) (+ 1 rz rh)  ))
                          (save (substring decs 0 (+ rz rh))) )
                      (if (< 5 (string-to-number overflow))
                          (progn
                            (setq decs (int-to-string (+ 1 (string-to-number save)))) ;; round up
                            (if (< (length decs) (length save)) (dotimes (i (- (length save) (length decs))) (setq decs (concat "0" decs)))  ;;cant drop leading zeros from decmil
                              (if (> (length decs) (length save))  ;; if round up went past decmil
                                  (let ((carry (substring decs 0 1)))
                                    (setq decs (substring decs 1 (1+ (length save))))
                                    (setq ints (int-to-string (+ (string-to-number carry) (string-to-number ints))))
                                    ))))
                        (setq decs save))))
                (if (< (length ints) lz) (dotimes (i (- lz (length ints))) (setq ints (concat "0" ints)))) ;; pad ones
                (if (< (length decs) rz) (dotimes (i (- lz (length decs))) (setq dec (concat decs "0" )))) ;; pad decmal
                (setq dec  (if (< 0 (length decs)) (concat ints "." decs) ints))

                )
              (if (string-match "," fmt)
                  (let ((re "^\\([0-9]+\\)\\(\\([0-9][0-9][0-9]\\)+\\(,[0-9][0-9][0-9]\\)*\\(\\.[0-9]*\\)?\\)$"))
                    (while (string-match re dec)
                      (setq dec (concat (match-string 1 dec) "," (match-string 2 dec)))
                      ) ) )
              (if exp
                  (setq dec (concat dec exp)) )
	      (if (> 0 value) 
		  (setq dec (cond		
			     ((string= "b" neg-type) (concat "(" dec ")"))
			     (t (concat "-" dec))
			     ))
		)
              dec)))      
      val)))



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


(defun qs-s-to-n (str) "converts a string to a number"
       (if (string-match "^ *\\([0123456789\\.,]+\\) ?\\([kMB%]\\|[eE][\\+-]?[0123456789]+\\)? *$" str)
           (let ((exp (match-string 2 str))
                 (num (string-to-number (string-replace "," "" (match-string 1 str))) )
                 (p 1))
             (if exp
                 (progn
                   (if (string-match "[eE]\\([\\+-]\\)?\\([0123456789]+\\)" exp)
                       (let ((sign (match-string 1 exp))
                             (power (match-string 2 exp))
                             (p 1))
                         (dotimes (i (string-to-number power))
                           (setq p (* 10 p)))
                         (if (equal "-" sign)
                             (setq p  (/ 1 (float p))))
                         (setq num (* p num)))
                     (setq p (cond
                              ((equal exp "%") .01)
                              ((equal exp "k") 1000)
                              ((equal exp "M") 1000000)
                              ((equal exp "B") 1000000000)
                              (t 1)
                              ))
                     )
                   (setq num (* p (float num)))
                   ))
             num)))



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
          ))
      )
    ))


(defun qs-eval-chain (addr chain)
  "Update cell and all deps"
  (let ((cc-addr (upcase (string-replace "$" "" addr)))
        (m (avl-tree-member qs-data (qs-addr-to-index addr))) )
    (if m
        (progn
          (qs-eval-fun addr)
          (dolist (a (elt m qs-c-deps))
            (if (not (listp chain)) (setq chain (list chain)))
            (if (not (member (upcase (string-replace "$" "" a)) chain))
                (qs-eval-chain a (append chain (list cc-addr)))
              )))
      )))


(defun qs-formula-cell-refs (formula)
  "Takes a string and finds all the cell references in a string Ex:
   (qs-fomual-cell-refs (\"= A1+B1\"))
   returns a list ((\"A1\" 2 3) (\"B1\" 5 6))"
  (let* ((s (concat formula " "))
         (j 0) (retval (list )))
    (while (string-match qs-cell-or-range-re s j)
      (setq retval (append retval (list (list (match-string 1 s) (match-beginning 1) (match-end 1)))))
      (setq j (match-end 1))
      )
    retval))

(defun qs-calc-eval-error (x) "stringify calc eval error message"
       (if (listp x) (concat "Error char " (number-to-string (elt x 0)) " :" (elt x 1)) x)
       )


(defun qs-eval-fmla (fmla)
  "evaluate fmla in calc and return value"
  (let* ((s (upper (string-replace "$" "" fmla)))
         (refs (qs-formula-cell-refs s))
         (cv "")
         )
    (dolist (ref refs)
      (setq cv (qs-cell-val (elt ref 0)))
      (setq s (concat (substring s 0 (elt ref  1)) cv (substring s (elt ref 2))))
      )
    ;;    (print (concat "Eval: " (substring s 1) " \\t to: " (calc-eval (substring s 1))))
    (setq cv (qs-calc-eval-error (calc-eval (substring s 1))))

    cv))

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
          (setq cv (qs-calc-eval-error (calc-eval (substring s 1))))

          (aset m qs-c-val cv)
          (aset m qs-c-fmtd (qs-format-number (elt m qs-c-fmt) cv))
          (let ((c 0)
                (i 0)
                (chra (- (string-to-char "A") 1)))
            (while (and  (< i (length addr)) (< chra (elt addr i)))
              (setq c (+ (* 26 c)  (- (logand -33 (elt addr i)) chra))) ;; -33 is mask to change case
              (setq i (+ 1 i)))
            (qs-draw-cell (- c 1) (string-to-number (substring addr i))
                          (qs-pad-right (elt m qs-c-fmtd) (- c 1))) )
          cv
          )
      (aref m qs-c-val)
      )))


(defun qs-add-dep (ca cc)
  "Add to dep list.
   CA -- source of information
   CC -- cell to update when ca changes"
  (let ((cc-addr (upcase (string-replace "$" "" cc)))
        (ca-addr (upcase (string-replace "$" "" ca))))
    (if (string-match "\\([A-Z]+[0-9]+\\):\\([A-Z]+[0-9]+\\)" ca-addr)
        (let* ((first-cell (qs-addr-to-rowcol (match-string 1 ca-addr)))
               (secnd-cell (qs-addr-to-rowcol (match-string 2 ca-addr)))
               (start-col (if (< (elt first-cell 0) (elt secnd-cell 0)) (elt first-cell 0) (elt secnd-cell 0)))
               (col-len (- (+ 1 (elt first-cell 0) (elt secnd-cell 0))  (* 2 start-col)))
               (start-row (if (< (elt first-cell 1) (elt secnd-cell 1)) (elt first-cell 1) (elt secnd-cell 1)))
               (row-len (- (+ 1 (elt first-cell 1) (elt secnd-cell 1))  (* 2 start-row)))
               )
          (dotimes (y row-len)
            (dotimes (x col-len)
              (qs-add-dep  (qs-rowcol-to-addr (+ x start-col) (+ y start-col)) cc-addr)
              )))
      (let  ( (m (avl-tree-member qs-data (qs-addr-to-index cc-addr))))
        (if m
            (let ((old-deps (elt m qs-c-deps)))
              (if (not (member cc old-deps) )
                  (aset m qs-c-deps (append (elt m qs-c-deps) (list ca-addr)))
                ))
          (progn
            (setq m (qs-new-cell cc-addr))
            (aset m qs-c-deps (list cc-addr))
            (avl-tree-enter qs-data m)
            ))))))


(defun qs-del-dep (ca cc)
  "ca - addr of cell with address
   cc - addr of cell to remove
   Remove cc from dep list of celll at ca."
  (if (and ca (not (string= "" ca)))
      (let ((cc-addr (upcase (string-replace "$" "" cc)))
            (ca-addr (upcase (string-replace "$" "" ca))))
        (if (string-match "\\([A-Z]+[0-9]+\\):\\([A-Z]+[0-9]+\\)" ca-addr)
            (let* ((first-cell (qs-addr-to-rowcol (match-string 1 ca-addr)))
                   (secnd-cell (qs-addr-to-rowcol (match-string 2 ca-addr)))
                   (start-col (if (< (elt first-cell 0) (elt secnd-cell 0)) (elt first-cell 0) (elt secnd-cell 0)))
                   (col-len (- (+ 1 (elt first-cell 0) (elt secnd-cell 0))  (* 2 start-col)))
                   (start-row (if (< (elt first-cell 1) (elt secnd-cell 1)) (elt first-cell 1) (elt secnd-cell 1)))
                   (row-len (- (+ 1 (elt first-cell 1) (elt secnd-cell 1))  (* 2 start-row)))
                   )
              (dotimes (y row-len)
                (dotimes (x col-len)
                  (qs-del-dep  (qs-rowcol-to-addr (+ x start-col) (+ y start-col)) cc-addr)
                  )))
          (let  ( (m (avl-tree-member qs-data (qs-addr-to-index ca-addr))))
            (if (and m (listp (aref m qs-c-deps)) (member cc-addr (aref m qs-c-deps)))
                (delete cc-addr (aref m qs-c-deps))
              )
            )
          )))
  )


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


(defun HELLO (x)
  (concat \"hello \" x )
  )

;; (defmath STOCK (stock)
;;   "gets Stock info"
;; (let (
;;       (urla "https://query1.finance.yahoo.com/v11/finance/quoteSummary/")
;;       (urlb "?modules=financialData,earnings") )
;;   (if (json-available-p)
;;       (json-parse-buffer (url-retrieve-synchronously (concat urla stock urlb) t t 15))
;;     nil )))

                                        ;       (with-current-buffer (url-retrieve-synchronously (concat urla stock urlb) t t q   15)
                                        ;       (buffer-string))

(provide 'qs-mode)
(push (cons "\\.xlsx$" 'qs-mode) auto-mode-alist)
(push (cons "\\.XLSX$" 'qs-mode) auto-mode-alist)
(push (cons "\\.csv$" 'qs-mode) auto-mode-alist)
(push (cons "\\.CSV$" 'qs-mode) auto-mode-alist)

;;; qs-mode.el ends here




;(defmath hello (name) "name" (interactive)
;  (concat "hello " name))


  

;;                                          _            _
;;        ___ ___       _ __ ___   ___   __| | ___   ___| |
;;       / __/ __|_____| '_ ` _ \ / _ \ / _` |/ _ \ / _ \ |
;;       \__ \__ \_____| | | | | | (_) | (_| |  __/|  __/ |
;;       |___/___/     |_| |_| |_|\___/ \__,_|\___(_)___|_|



;;; ss-mode -- Spreadsheet Mode -- Tabular interface to Calc
;; Copyright (C) 2014 -- Use at 'yer own risk  -- NO WARRANTY!
;; Author: sam fiechter sam.fiechter(at)gmail
;; Version: 0.000000000000001
;; Created: 2014-03-24
;; Keywords: calc, spreadsheet


;;;Code
(require 'avl-tree)
;; ;;;;;;;;;;;;; variables ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defvar ss-empty-name "*Sams Spreadsheet Mode*")

(defvar ss-cur-col 0)
(defvar ss-max-col 3)
(defvar ss-col-widths (make-vector ss-max-col 7))

(defvar ss-cur-row 1)
(defvar ss-max-row 3)
(defvar ss-row-padding 4)
(defvar ss-data (avl-tree-create 'ss-avl-cmp))
(defvar ss-input-history (list ))
;;       CELL Format -- [ "A1" 0.5 "= 1/2" "%0.2g" "= 3 /2" (list of cells to calc when changes)]
;;   0 = index / cell Name
(defvar ss-c-addr 0)
;;   1 = value (number)
(defvar ss-c-val 1)
;;   2 = format (TBD)
(defvar ss-c-fmt 2)
;;   3 = formula
(defvar ss-c-fmla 3)
;;   4 = depends on -- list of indexes
(defvar ss-c-deps 4)
;; ;;;;;;;;;;;;; keymaps ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defvar ss-map  (make-sparse-keymap 'ss-map))


(define-key ss-map [left]       'ss-move-left)
(define-key ss-map [right]      'ss-move-right)
(define-key ss-map [tab]        'ss-move-right)
(define-key ss-map [up]         'ss-move-up)
(define-key ss-map [down]       'ss-move-down)
(define-key ss-map [return]     'ss-edit-cell)
(define-key ss-map [backspace]  'ss-clear-key)
(define-key ss-map [remap self-insert-command] 'ss-edit-cell)


;;(define-key ss-map [C-R]        'ss-search-buffer)
                                        ;(define-key ss-map (kbd "RET")        'ss-edit-cell)
(define-key ss-map (kbd "+")  'ss-increase-cur-col )
(define-key ss-map (kbd "-")  'ss-decrease-cur-col )

;;  _   _                 ___       _             __
;; | | | |___  ___ _ __  |_ _|_ __ | |_ ___ _ __ / _| __ _  ___ ___
;; | | | / __|/ _ \ '__|  | || '_ \| __/ _ \ '__| |_ / _` |/ __/ _ \
;; | |_| \__ \  __/ |     | || | | | ||  __/ |  |  _| (_| | (_|  __/
;;  \___/|___/\___|_|    |___|_| |_|\__\___|_|  |_|  \__,_|\___\___|




(defun ss-increase-cur-col ()
  "Increase the width of the current column" (interactive)
  (aset ss-col-widths ss-cur-col (+ 1 (elt ss-col-widths ss-cur-col)))
  (ss-draw-all))

(defun ss-decrease-cur-col ()
  "Decrease the width of the current column" (interactive)
  (aset ss-col-widths ss-cur-col (- (elt ss-col-widths ss-cur-col) 1))
  (ss-draw-all))


(defun ss-input-cursor-move (dir)
  "Move the cursor" (interactive)
  (if (and (>= 0 (+ dir ss-input-cursor)) (<= (+ ss-input-cursor dir) (length ss-input-buffer)))
      (setq ss-input-cursor (+ ss-input-cursor dir))
    nil
    ))

(defun ss-buffer-rot (dir)
  "Rotate the Command Buffer"
  (message "do buffer rot")
  )

(defun ss-backspace-key ()
  "Delete from buffer"  (interactive)
  (if (and (> 0 (length ss-input-buffer)) (0 > ss-input-cursor))
      (let ((nc (- ss-input-cursor -1)))  ;; 01234
        (setq ss-input-buffer (concat (substring ss-input-buffer 0 nc) (substring ss-input-cursor (+ 1 ss-input-cursor))))
        (setq ss-input-cursor nc)
        )))

;;  ____        _           ___       _     _      	
;; |  _ \  __ _| |_ __ _   / / \   __| | __| |_ __ 	
;; | | | |/ _` | __/ _` | / / _ \ / _` |/ _` | '__|	
;; | |_| | (_| | || (_| |/ / ___ \ (_| | (_| | |   	
;; |____/ \__,_|\__\__,_/_/_/   \_\__,_|\__,_|_|    functions
;; 

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
      
	
  


(defun ss-avl-cmp (a b)
  (let ((A (if (sequencep a) (elt a ss-c-addr) a)) (B (if (sequencep b) (elt b ss-c-addr) b)))  ; a or b can be vectors or addresses
    (< (ss-addr-to-index A) (ss-addr-to-index B)) ))


(defun ss-addr-to-index (a)
  "Convert from ss addr (e.g. A1) to index  -- expects [A-Z]+[0-9]+"
  (let ((chra (- (string-to-char "A") 1)))
    (if (sequencep a)
        (progn
          (let ((out 0) (i 0))
            (while (and  (< i (length a)) (< chra (elt a i)))
              (setq out (+ (* 26 out)  (- (logand -33 (elt a i)) chra))) ;; -33 is mask to change case
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

(defun ss-col-letter (a)
  "returns the letter of the column id arg -- expecst int"
  (let ((out "") (n 1) (chra (string-to-char "A")))
    (while (<= 0 a)
      (setq out (concat (char-to-string (+ chra (% a 26))) out))
      (setq a  (- (/ a 26) 1)) )
    out ))

;;  ____                     _
;; |  _ \ _ __ __ ___      _(_)_ __   __ _
;; | | | | '__/ _` \ \ /\ / / | '_ \ / _` |
;; | |_| | | | (_| |\ V  V /| | | | | (_| |
;; |____/|_|  \__,_| \_/\_/ |_|_| |_|\__, |
;;                                   |___/
;; functions dealing with the cursor and cell drawing /padding

(defun ss-highlight (txt)
  "highlight text"
  (let ((myface '((:bold t) (:foreground "White") (:background "Blue") (:invert t))))
  (propertize txt 'font-lock-face myface)))
;;  (concat ">" txt "<"))

(defun ss-pad-center (s i)
  "pad a string out to center it - expects stirng, col no (int)"
  (let ( (ll (length s) ))
    (if (>= ll (elt ss-col-widths i))
        (make-string (elt ss-col-widths i) (string-to-char "#"))
      (let* ( (pl (/ (- (* 2 (elt ss-col-widths i)) ll)  2))  ; half the pad length
              (pad (make-string (- (elt ss-col-widths i) (+ pl ll)) (string-to-char " "))) )
        (concat (make-string pl (string-to-char " ")) s pad)) )
    ))

(defun ss-pad-left (s i)
  "pad to the left  - expects stirng, col no (int)"
  (let ( (ll (length s)) )
    (if (>= ll (elt ss-col-widths i))
        (make-string (elt ss-col-widths i) (string-to-char "#"))
      (concat s (make-string (- (elt ss-col-widths i)  ll) (string-to-char " ")))
      )))

(defun ss-pad-right (s i)
  "pad to the right  - expects stirng, col no (int)"
  (let ( (ll (length s)) )
    (if (>= ll (elt ss-col-widths i))
        (make-string (elt ss-col-widths i) (string-to-char "#"))
      (concat (make-string (- (elt ss-col-widths i)  ll) (string-to-char " ")) s)
      )))



(defun ss-draw-all ()
  "Populate the current ss buffer." (interactive)
  (pop-to-buffer ss-empty-name nil)
  (let ((i 0) (j 0) (k 0) (header (make-string ss-row-padding (string-to-char " "))))
    (beginning-of-buffer)
    (erase-buffer)
    (dotimes (i ss-max-col)  ;; make header
      (setq header (concat header (ss-pad-center (ss-col-letter i) i))))
    (dotimes (j ss-max-row) ;; draw buffer
      (if (= 0 j)
          (insert  (ss-highlight header))
        (progn
          (insert (ss-highlight (format (concat "%" (int-to-string ss-row-padding) "d") j) ))
          (dotimes (i ss-max-col)
            (let ((m (avl-tree-member ss-data (+ j (* ss-max-row (+ i 1))))))
              (if m
                  (insert (ss-pad-right (elt m ss-c-val) i))
                (insert (make-string (elt ss-col-widths i) (string-to-char " "))))))))

      (insert "\n"))
    (set-buffer-modified-p nil)
    (ss-move-cur-cell 0 0) ;;draw cursor
    ))


(defun ss-draw-cell (x y text)
  "redraw one cell on the ss  - expects  col no. (int), row no, (int), cell value (padded string)"
  (let ((col ss-row-padding) (i 0))
    (pop-to-buffer ss-empty-name nil)
    (setq cursor-type nil)  ;; no cursor
    (setq truncate-lines 1)  ;; no wrap-around
    (dotimes (i x)
      (setq col (+ col (elt ss-col-widths i))) )
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

  (if (< ss-cur-col (- ss-max-col 1))
      (ss-move-cur-cell 1 0)
    (progn
      (setq ss-max-col (+ 1 ss-max-col))
      (setq ss-col-widths   (vconcat ss-col-widths (list (elt ss-col-widths ss-cur-col))))
      (setq ss-cur-col (+ 1 ss-cur-col))
      (ss-draw-all)
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
      )))

(defun ss-move-cur-cell (x y) (interactive)
       (let* ((new-row (+ ss-cur-row y))
              (new-col (+ ss-cur-col x))
              (m (avl-tree-member ss-data (+ ss-cur-row (* ss-max-row (+ ss-cur-col 1)))))
              (n (avl-tree-member ss-data (+ new-row (* ss-max-row (+ new-col 1)))))
              (ot (if m (elt m ss-c-val) ""))
              (nt (if n (elt n ss-c-val) "")))
         (progn
           (ss-draw-cell ss-cur-col ss-cur-row  (ss-pad-right ot ss-cur-col))
           (setq ss-cur-row new-row)
           (setq ss-cur-col new-col)

           (ss-draw-cell ss-cur-col ss-cur-row  (ss-highlight  (ss-pad-right nt ss-cur-col) ))
	   (message (concat "Cell " (ss-col-letter ss-cur-col) (int-to-string ss-cur-row) " : " (if n (if (string= "" (elt n ss-c-fmla)) (elt n ss-c-val ) (elt n ss-c-fmla)) "" )))
	   )))


;;   ____     _ _   _____            _             _   _
;;  / ___|___| | | | ____|_   ____ _| |_   _  __ _| |_(_) ___  _ __
;; | |   / _ \ | | |  _| \ \ / / _` | | | | |/ _` | __| |/ _ \| '_ \
;; | |__|  __/ | | | |___ \ V / (_| | | |_| | (_| | |_| | (_) | | | |
;;  \____\___|_|_| |_____| \_/ \__,_|_|\__,_|\__,_|\__|_|\___/|_| |_|
;; fuctions dealing with eval


(defun ss-cell-val (addr)
  "get the value of a cell"
  (let ((m (avl-tree-member ss-data  (ss-addr-to-index addr)) ))
    (if m (elt m ss-c-val) "0")
    ))

(defun ss-eval-chain (addr chain)
  "Updaet cell and all deps"
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
         (re "[^A-Za-z0-9]\\([A-Za-z]+[0-9]+\\)[^A-Za-z0-9\(]")
         (j nil) (retval (list )))
    (if (string-match re s )
        (progn
          (while (not (equal j (match-end 1)))
            (setq retval (append retval (list (list (substring s (match-beginning 1) (match-end 1)) (match-beginning 1) (match-end 1)))))
            (setq j (match-end 1))
            (string-match re s j) )) nil )
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
          (let ((c 0) (i 0) (chra (- (string-to-char "A") 1)))
            (while (and  (< i (length addr)) (< chra (elt addr i)))
              (setq c (+ (* 26 c)  (- (logand -33 (elt addr i)) chra))) ;; -33 is mask to change case
              (setq i (+ 1 i)))
            (ss-draw-cell (- c 1) (string-to-int (substring addr i)) (ss-pad-right cv (- c 1))) )
          cv ) (aref m ss-c-val)
          )))


(defun ss-add-dep (ca cc)
  "Add to dep list."
  (let  ( (m (avl-tree-member ss-data (ss-addr-to-index ca)) ))
    (if m
        (aset m ss-c-deps (append (elt m ss-c-deps) (list cc)))
      (progn
        (setq m (vector ca "" "" "" (list cc)))
        (avl-tree-enter ss-data m)
        ))))

(defun ss-del-dep (ca cc)
  "Remove from dep list."
  (let  ( (m (avl-tree-member ss-data (ss-addr-to-index ca))))
    (if m
        (delete cc (aref m ss-c-deps))
      nil) ))

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
    ;;delete this cell from its old deps

    (if m (let ((refs (ss-formula-cell-refs  (aref m ss-c-fmla))))
            (dolist (ref refs)
              (ss-del-dep (elt ref 1) current-cell))
            ))

    ;; if this is a formula, do deps
    (if (= (string-to-char "=") (elt nt 0)) ; new value starts with  =
        (progn
          (if m
              (aset m ss-c-fmla nt)
            (progn
              (setq m (vector current-cell "0" "0" nt (list)))
              (avl-tree-enter ss-data m)
              ))
          (let*  ((cv "") (ca "")
                  (s (elt m ss-c-fmla))
                  (delta 0)
                  (refs (ss-formula-cell-refs s)) )
            (if refs
                (progn
                  (dolist (ref refs)
                    (setq ca (elt ref 0))
                    (ss-add-dep ca current-cell)
                    (setq cv (ss-cell-val ca))
                    (setq s (concat (substring s 0 (+ delta (elt ref  1))) cv (substring s (+ (elt ref 2) delta ))))
                    (setq delta (+ delta (- (length cv) (length ca))))
                    )))
            (setq nt (calc-eval (substring s 1)))

            )) nil )
    ;; if not a formula
    (if m (aset m ss-c-val nt)
      (progn  ;;else
        (setq m (vector current-cell nt "0" "" (list)))
        (avl-tree-enter ss-data m)  ))

    ;;do eval-chain
    (dolist (a (aref m ss-c-deps))
      (if (not (equal a current-cell))  (ss-eval-chain a current-cell)  nil))

    ;;draw it
    (ss-draw-cell ss-cur-col ss-cur-row (ss-highlight (ss-pad-right nt ss-cur-col) ))
    )
    (ss-move-down)
    )


(defun ss-close () (interactive)
       (kill-buffer (current-buffer)) )

;; (defun ss-import-csv (filename)
;;   "Read a CSV file into ss-mode"
;;   (interactive "f")
;;   (find-file (concat filename ".ses"))
;;   (setq ss-mode-data [ ])
;;   (with-temp-buffer
;;    (insert-file-contents filename)
;;      (beginning-of-buffer)
;;      (while (not (eobp))
;;        (let ((x 0) (cell "") (row [])
;;           (line (thing-at-point 'line)))
;;       (dolist (cell (split-string line ","))
;;         (if (and (char-equal "\"" (substring cell 0 0))
;;                  (char-equal "\"" (substring cell -1 -1)))
;;             (setq cell (substring cell 1 -1)) nil )
;;         (setq row (vconcat row (list  cell))) )
;;       (setq ss-mode-data (vconcat ss-mode-data row))
;;       (with-current-buffer (ses-goto-print (x+1) 0))
;;       (next-line)
;;       ))) )



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


  (pop-to-buffer ss-empty-name nil)

  (setq ss-cur-col 0)
  (setq ss-max-col 3)
  (setq ss-col-widths (make-vector ss-max-col 7))
  (setq ss-cur-row 1)
  (setq ss-max-row 3)
  (setq ss-row-padding 4)
  (setq ss-data (avl-tree-create 'ss-avl-cmp))

  
  (setq ss-input-stack (list))
  (setq ss-input-stack-idx 0)
  (setq ss-input-buffer "")
  (setq ss-input-cursor 0)




  
  (ss-draw-all)

)

;;;###autoload
(defun ss-mode-open ()
  "Open SS mode
     ss keybindings:
     \\<ss-mode-map>
\\[ss-start-game]        Start a new game
\\[ss-end-game]        Terminate the current game
\\[ss-move-left]        Moves the board to the left
\\[ss-move-right]        Moves the board to the right
\\[ss-move-up]        Moves the board to the up
\\[ss-move-down]        Moves the board to the down
"
  (interactive)
  (pop-to-buffer ss-empty-name nil)
  (ss-mode) )



(provide 'ss-mode)

;;; ss-mode.el ends here

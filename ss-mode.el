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
(defvar ss-range-parts-re "\\([A-Za-z]+\\)\\([0-9]+\\)\\:\\([A-Za-z]+\\)\\([0-9]+\\)")
(defvar ss-one-cell-re "$?[A-Za-z]+$?[0-9]+")
(defvar ss-range-re "$?[A-Za-z]+$?[0-9]+\\:$?[A-Za-z]+$?[0-9]+")
(defvar ss-cell-or-range-re (concat "[^A-Za-z0-9]?\\(" ss-range-re  "\\|" ss-one-cell-re  "\\)[^A-Za-z0-9\(]?"))

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
  "This is the function used by avl tree to compare ss addresses"
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
  "returns the letter of the column id arg -- expects int"
  (let ((out "") (n 1) (chra (string-to-char "A")))
    (while (<= 0 a)
      (setq out (concat (char-to-string (+ chra (% a 26))) out))
      (setq a  (- (/ a 26) 1)) )
    out ))


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
    (if (= (string-to-char "=") (elt nt 0)) ; formulas start with  =
        (progn
          (if m
              (aset m ss-c-fmla nt)
            (progn
              (setq m (vector current-cell "0" "0" nt (list)))
              (avl-tree-enter ss-data m)
              ))

	  (let ((s (concat nt " ")) (j 0) (d 0) (ms "") (cv "") (me 0) (mt 0))
	    (while (and (< j (length s)) (string-match ss-cell-or-range-re s j))
	      (setq ms (match-string 1 s))
	      (setq me (match-end 1))
	      (setq mt (match-beginning 1))
	      (setq cv (ss-cell-val ms))
	      (ss-add-dep ms current-cell)
	      (setq d (+ me (- (length cv) (length ms))))
	      (setq s (concat (if (< 0 mt) (substring s 0 mt ) "")
			      cv
			      (if (< me (length s)) (substring s me) "" )
			      ))
	      (setq j d))
            (setq nt (calc-eval (substring s 1 -1)))  ;; throw away = and  ' '
           )) nil )
      ;; nt is now not a formula
      (if m (aset m ss-c-val nt)
	(progn  ;;else
	  (setq m (vector current-cell nt "0" "" (list)))
	  (avl-tree-enter ss-data m)
	  ))
      ;;do eval-chain
    (dolist (a (aref m ss-c-deps))
      (if (equal a current-cell)
	  nil ;; don't loop infinite
	(ss-eval-chain a current-cell)))

    ;;draw it
    (ss-draw-cell ss-cur-col ss-cur-row (ss-highlight (ss-pad-right nt ss-cur-col) ))
    ))

(defun ss-import-csv (filename)
  "Read a CSV file into ss-mode"
  (interactive "fFilename:")
  (let ((mybuff (current-buffer))
	(cc ss-cur-col) 
	(cr ss-cur-row))
    
  (with-temp-buffers
   (insert-file-contents filename)
   (beginning-of-buffer)
   (let ((x 0) (cell "") (cl ss-cur-col)
	 (rw ss-cur-row) (j 0)
	 (re ",?\"\\([^\"]+\\)\"[\n,]\\|,?\\([^,]+\\)[\n,]")
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

(defun ss-highlight (txt)
  "highlight text"

  (let ((myface '((:foreground "White") (:background "Blue"))))
  (propertize txt 'font-lock-face myface)))
;;  (concat ">" txt "<"))

(defun ss-pad-center (s i)
  "pad a string out to center it - expects stirng, col no (int)"
  (setq s (replace-regexp-in-string "\n" " " s))
  (let ( (ll (length s) ))
    (if (>= ll (elt ss-col-widths i))
        (make-string (elt ss-col-widths i) (string-to-char "#"))
      (let* ( (pl (/ (- (* 2 (elt ss-col-widths i)) ll)  2))  ; half the pad length
              (pad (make-string (- (elt ss-col-widths i) (+ pl ll)) (string-to-char " "))) )
        (concat (make-string pl (string-to-char " ")) s pad)) )
    ))

(defun ss-pad-left (s i)
  "pad to the left  - expects stirng, col no (int)"
  (setq s (replace-regexp-in-string "\n" " " s))
  (let ( (ll (length s)) )
    (if (>= ll (elt ss-col-widths i))
        (make-string (elt ss-col-widths i) (string-to-char "#"))
      (concat s (make-string (- (elt ss-col-widths i)  ll) (string-to-char " ")))
      )))

(defun ss-pad-right (s i)
  "pad to the right  - expects stirng, col no (int)"
  (setq s (replace-regexp-in-string "\n" " " s))
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
	   (minibuffer-message (concat "Cell " (ss-col-letter ss-cur-col) (int-to-string ss-cur-row)
				       " : " (if n (if (string= "" (elt n ss-c-fmla)) (elt n ss-c-val ) (elt n ss-c-fmla)) "" )))
	   )))


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
          (let ((c 0) (i 0) (chra (- (string-to-char "A") 1)))
            (while (and  (< i (length addr)) (< chra (elt addr i)))
              (setq c (+ (* 26 c)  (- (logand -33 (elt addr i)) chra))) ;; -33 is mask to change case
              (setq i (+ 1 i)))
            (ss-draw-cell (- c 1) (string-to-int (substring addr i)) (ss-pad-right cv (- c 1))) )
          cv ) (aref m ss-c-val)
          )))


(defun ss-add-dep (ca cc)
  "Add to dep list."
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
	  (setq m (vector addr "" "" "" (list cc)))
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

	;;  ____        __ __  __       _   _         	
	;; |  _ \  ___ / _|  \/  | __ _| |_| |__  ___ 	
	;; | | | |/ _ \ |_| |\/| |/ _` | __| '_ \/ __|	
	;; | |_| |  __/  _| |  | | (_| | |_| | | \__ \	
	;; |____/ \___|_| |_|  |_|\__,_|\__|_| |_|___/	



(defmath sum (x)
  "add the items in the range"
  (interactive 1 "sum")
  :" vflat(x) * ((vflat(x) *0) +1)"
)	
  

(provide 'ss-mode)

;;; ss-mode.el ends here

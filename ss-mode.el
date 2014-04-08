;	                                   _            _ 	
;	 ___ ___       _ __ ___   ___   __| | ___   ___| |	
;	/ __/ __|_____| '_ ` _ \ / _ \ / _` |/ _ \ / _ \ |	
;	\__ \__ \_____| | | | | | (_) | (_| |  __/|  __/ |	
;	|___/___/     |_| |_| |_|\___/ \__,_|\___(_)___|_|	
;	                                                  	

;;; ss-mode -- Spreadsheet Mode -- Tabular interface to Calc
;; Copyright (C) 2014 -- Use at 'yer own risk  -- NO WARRANTY!
;; Author: sam fiechter sam.fiechter(at)gmailxs
;; Version: 0.000000000000001
;; Created: 2014-03-24
;; Keywords: calc, tabular 


;;;Code
(require 'avl-tree)
;; ;;;;;;;;;;;;; variables ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defface ss-mode-highlight-face   '((((class grayscale) (background light))
				     (:background "Black" :foreground "White" :bold t))
				    (((class grayscale) (background dark))
				     (:background "white" :foreground "black" :bold t)) 
				    (((class color) (background light))
				     (:background "Blue" :foreground "White" :bold t))
				    (((class color) (background dark))
				     (:background "White" :foreground "Blue" :bold t))
				    (t (:inverse-video t :bold t)))
  "ss-mode face used to highlight cells"
  :group 'ss-mode)

(defvar ss-mode-empty-name "*Sams Spreadsheet Mode*")

(defvar ss-mode-cur-col 0)
(defvar ss-mode-max-col 10)
(defvar ss-mode-col-widths (make-vector ss-mode-max-col 7))


(Defvar ss-mode-cur-row 0)

(defvar ss-mode-max-row 10)
(defvar ss-mode-data (avl-tree-create 'ss-mode-avl-cmp))

;; ;;;;;;;;;;;;; keymaps ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defvar ss-mode-map  (make-sparse-keymap 'ss-mode-map))

(define-key ss-mode-map [left]          'ss-move-left)
(define-key ss-mode-map [right]        'ss-move-right)
(define-key ss-mode-map [up]           'ss-move-up)
(define-key ss-mode-map [down]         'ss-move-down)
(define-key ss-mode-map (kbd "RET")        'ss-edit-cell)

;; ;;;;;;;;;;;;; functions ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; SAM's SES MACROS

(defun ss-mode-avl-cmp (a b)
  (let ((A (if (sequencep a) (elt a 0) a)) (B (if (sequencep b) (elt b 0) b)))  ; a or b can be vectors or addresses
    (< (ss-mode-to-index A) (ss-mode-to-index B)) ) )
       
(defun ss-mode-to-index (a)
  "Convert from ss-mode to index  -- requires [A-Z]+[0-9]+"
  (if (sequencep a)
      (progn 
	(let ((out 0) (i 0))
	  (while (and  (< i (length a)) (< 64 (elt a i)))
	    (setq out (+ (* 26 out)  (- (logand -33 (elt a i)) 64)))
	    (setq i (+ 1 i)))
	  (setq out (+ (* out ss-mode-max-row) (string-to-int (substring a i))))
	  out))
    a) )

(defun ss-mode-col-letter (a)
  "returns the letter of the column id arg"
  (let ((out "") (n 1))
    (while (<= 0 a)
      (setq out (concat (char-to-string (+ 65 (% a 26))) out))
      (setq a  (- (/ a 26) 1)) )
    out ))

(defun ss-pad-center (s i)
  "pad a string out to center it"
  (let* ( (ll (length s))
	  (pl (/ (- (* 2 (elt ss-mode-column-widths i) ll)  2)))  ; half the pad length
	  (pad (make-string (- (elt ss-mode-column-widths i) (+ pl ll)) " " )))
	(concat (make-string pl " ") s pad) ))

(defun ss-pad-left (s i)
  "pad to the leftt"
  (let ( (ll (length s))
	 (pad (make-string (- (elt ss-mode-column-widths i) (+ pl ll)) " " )))
	(concat s pad) ))

(defun ss-pad-right (s i)
  "pad to the right"
  (let* ( (ll (length s))
	  (pad (make-string (- (elt ss-mode-column-widths i) (+ pl ll)) " " )))
	(concat pad s) ))


(defun ss-mode-print ()
  "Populate the current ss-mode buffer."
  (let ((i 0) (j 0) (k 0))
    (erase-buffer)
    (let ((header "    "))
    (dotimes (i ss-mode-max-col)
      (setq header (concat header (ss-pad-center (ss-mode-col-letter i) i))))
    (dotimes (j ss-mode-max-row)
      (if (= 0 j)
	  (insert  (propertize header 'font-lock-face '(:inverse-video t)))
	(insert (propertize (ss-pad-center (int-to-string j) j) 'font-lock-face '(:inverse-video t)))
	(dotimes (i ss-mode-max-col)
	  (move-end-of-line)
	(let ((m (avl-tree-member (+ j (* ss-mode-max-row (+ i 1))))))
	      (if (m)
		  (insert (ss-pad-right (number-to-string (elt m 1) i)))
		(insert (make-string (elt ss-mode-column-widths i) " ")) ) ) ) ))
    (set-buffer-modified-p nil)

    ;; If REMEMBER-POS was specified, move to the "old" location.
    (if saved-pt
	(progn (goto-char saved-pt)
	       (move-to-column saved-col)
	       (recenter))
      (goto-char (point-min)))))

(defun tabulated-list-print-entry (id cols)
  "Insert a Tabulated List entry at point.
This is the default `tabulated-list-printer' function.  ID is a
Lisp object identifying the entry to print, and COLS is a vector
of column descriptors."
  (let ((beg   (point))
	(x     (max tabulated-list-padding 0))
	(ncols (length tabulated-list-format))
	(inhibit-read-only t))
    (if (> tabulated-list-padding 0)
	(insert (make-string x ?\s)))
    (dotimes (n ncols)
      (setq x (tabulated-list-print-col n (aref cols n) x)))
    (insert ?\n)
    (put-text-property beg (point) 'tabulated-list-id id)
    (put-text-property beg (point) 'tabulated-list-entry cols)))


(defun ss-move-left ()
  (interactive)
  (ss-mode-move-cur-cell -1 0))

(defun ss-move-right ()
  (interactive)
  (ss-mode-move-cur-cell 1 0))

(defun ss-move-up ()
  (interactive)
  (ss-mode-move-cur-cell 0 -1))

(defun ss-move-down ()
  (interactive)
  (ss-mode-move-cur-cell 0 1 ))

;;;###autoload
(define-derived-mode ss-mode text-mode ss-mode-empty-name
  "ss game mode
  Keybindings:
  \\{ss-mode-map} "
  (use-local-map ss-mode-map)
 
 ;; (unless (featurep 'emacs)
  ;;   (setq mode-popup-menu
  ;;         '("ss Commands"
  ;;           ["Start new game"        ss-start-game]
  ;;           ["End game"                ss-end-game
  ;;            (ss-active-p)]
  ;;           ))


      (pop-to-buffer ss-mode-empty-name nil)
  (setq cursor-type nil)
  (setq ss-mode-cur-row 0)
  (setq ss-mode-cur-col 0)
  (ss-mode-move-cur-cell 0 0 )
  ) 

;;;###autoload
(defun ss-mode-fun ()
  "Open SS mode
     ss-mode keybindings:
     \\<ss-mode-map>
\\[ss-start-game]        Start a new game
\\[ss-end-game]        Terminate the current game
\\[ss-move-left]        Moves the board to the left
\\[ss-move-right]        Moves the board to the right
\\[ss-move-up]        Moves the board to the up
\\[ss-move-down]        Moves the board to the down
"
  (interactive)
  (pop-to-buffer ss-mode-empty-name nil)
  (ss-mode) )




(defun ss-mode-move-cur-cell (x y) (interactive)
  (let* ((row (elt (elt tabulated-list-entries ss-mode-cur-row) 1))
	 (cell-value (aref row ss-mode-cur-col))
	 (new-row (+ ss-mode-cur-row y))
	 (new-col (+ ss-mode-cur-col x)))
	 (if (and (<= 0 new-row) (> (length tabulated-list-entries) new-row)
		  (<= 0 new-col) (> (length row) new-col))
	     (progn
	       (aset row ss-mode-cur-col (propertize cell-value 'font-lock-face '(:default t)))
	       (setq ss-mode-cur-row new-row)
	       (setq ss-mode-cur-col new-col)
	       (setq row (elt (elt tabulated-list-entries ss-mode-cur-row) 1))
	       (setq cell-value (aref row ss-mode-cur-col))
	       (aset row ss-mode-cur-col (propertize cell-value 'font-lock-face '(:inverse-video t))) )))
  (tabulated-list-print t) )
  

(defun ss-edit-cell ()
(interactive)
(let* ((row (elt (elt tabulated-list-entries ss-mode-cur-row) 1))
       (cell-value (read-string "Cell Value:" (aref row ss-mode-cur-col))))
  (if (= 61 (elt cell-value 0)) ; new value starts with =
      (ss-set-function cell-value)
     (aset row ss-mode-cur-col (propertize cell-value 'font-lock-face '(:inverse-video t))) ))
  (tabulated-list-print t) )

(defun ss-set-function (fun)
  "Sets the function for current cell to value"

  )



(defun ss-close () (interactive)
  (kill-buffer (current-buffer)) )


(defun ss-import-csv (filename)
  "Read a CSV file into ss-mode"
  (interactive "f")
  (find-file (concat filename ".ses"))
  (setq ss-mode-data [ ])
  (with-temp-buffer
   (insert-file-contents filename)
     (beginning-of-buffer)
     (while (not (eobp))
       (let ((x 0) (cell "") (row [])
	     (line (thing-at-point 'line)))
	 (dolist (cell (split-string line ","))
	   (if (and (char-equal "\"" (substring cell 0 0))
		    (char-equal "\"" (substring cell -1 -1)))
	       (setq cell (substring cell 1 -1)) nil )
	   (setq row (vconcat row (list  cell))) )
	 (setq ss-mode-data (vconcat ss-mode-data row))
	 (with-current-buffer (ses-goto-print (x+1) 0))
	 (next-line) 
	 ))) )





(provide 'ss-mode)

;;; ss-mode.el ends here

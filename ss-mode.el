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
;; Keywords: calc, spreadsheet


;;;Code
(require 'avl-tree)
;; ;;;;;;;;;;;;; variables ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defface ss-highlight-face   '((((class grayscale) (background light))
				     (:background "Black" :foreground "White" :bold t))
				    (((class grayscale) (background dark))
				     (:background "white" :foreground "black" :bold t)) 
				    (((class color) (background light))
				     (:background "Blue" :foreground "White" :bold t))
				    (((class color) (background dark))
				     (:background "White" :foreground "Blue" :bold t))
				    (t (:inverse-video t :bold t)))
  "ss- face used to highlight cells"
  :group 'ss-)

(defvar ss-empty-name "*Sams Spreadsheet Mode*")

(defvar ss-cur-col 0)
(defvar ss-max-col 3)
(defvar ss-col-widths (make-vector ss-max-col 7))

(defvar ss-cur-row 1)
(defvar ss-max-row 3)
(defvar ss-row-padding 4)
(defvar ss-data (avl-tree-create 'ss-avl-cmp))

;	AVL Format -- [ "A1" 0.5 "= 1/2" "%0.2g" (list of cells to calc when changes)
;



;; ;;;;;;;;;;;;; keymaps ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defvar ss-map  (make-sparse-keymap 'ss-map))

(define-key ss-map [left]          'ss-move-left)
(define-key ss-map [right]        'ss-move-right)
(define-key ss-map [up]           'ss-move-up)
(define-key ss-map [down]         'ss-move-down)
(define-key ss-map (kbd "RET")        'ss-edit-cell)

;; ;;;;;;;;;;;;; functions ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; SAM's SES MACROS

(defun ss-avl-cmp (a b)
  (let ((A (if (sequencep a) (elt a 0) a)) (B (if (sequencep b) (elt b 0) b)))  ; a or b can be vectors or addresses
    (< (ss-to-index A) (ss-to-index B)) ))
       
(defun ss-to-index (a)
  "Convert from ss to index  -- requires [A-Z]+[0-9]+"
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

(defun ss-col-letter (a)
  "returns the letter of the column id arg"
  (let ((out "") (n 1) (chra (string-to-char "A")))
    (while (<= 0 a)
      (setq out (concat (char-to-string (+ chra (% a 26))) out))
      (setq a  (- (/ a 26) 1)) )
    out ))

(defun ss-pad-center (s i)
  "pad a string out to center it"
  (let* ( (ll (length s))
	  (pl (/ (- (* 2 (elt ss-col-widths i)) ll)  2))  ; half the pad length
	  (pad (make-string (- (elt ss-col-widths i) (+ pl ll)) (string-to-char " "))) )
	(concat (make-string pl (string-to-char " ")) s pad) ))

(defun ss-pad-left (s i)
  "pad to the leftt"
  (let ( (ll (length s))
	 (pad (make-string (- (elt ss-col-widths i)  ll) (string-to-char " ") )))
	(concat s pad) ))

(defun ss-pad-right (s i)
  "pad to the right"
  (let* ( (ll (length s))
	  (pad (make-string (- (elt ss-col-widths i) ll) (string-to-char " ") )))
	(concat pad s) ))


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
	  (insert  (propertize header 'font-lock-face '(:inverse-video t)))
	(progn	  
	  (insert (propertize (format (concat "%" (int-to-string ss-row-padding) "d") j) 'font-lock-face '(:inverse-video t)))
	  (dotimes (i ss-max-col)
	    (let ((m (avl-tree-member ss-data (+ j (* ss-max-row (+ i 1))))))
	      (if m
		  (insert (ss-pad-right (elt m 1) i))
		(insert (make-string (elt ss-col-widths i) (string-to-char " "))))))))

      (insert "\n"))
    (set-buffer-modified-p nil)
    (ss-move-cur-cell 0 0))) ;;draw cursor
    

(defun ss-draw-cell (x y text)
  "redraw one cell on the ss"
  (let ((col ss-row-padding) (i 0))
    (pop-to-buffer ss-empty-name nil)
    (dotimes (i x)
      (setq col (+ col (elt ss-col-widths i))) )
    (goto-line (+ y 1))
    (move-to-column col)
    (message (format "Going to %d, %d" col (+ y 1)))
    (insert text)
    (delete-forward-char (length text))
    (recenter) ))


(defun ss-move-left ()
  (interactive)
  (if (> ss-cur-col  0)
      (ss-move-cur-cell -1 0) nil) )

(defun ss-move-right ()
  (interactive)

  (if (< ss-cur-col (- ss-max-col 1))
      (ss-move-cur-cell 1 0)
    (progn
      (setq ss-max-col (+ 1 ss-max-col))
      (setq ss-col-widths (vconcat ss-col-widths (elt ss-col-widths ss-cur-col)))
      (setq ss-cur-col (+ 1 ss-cur-col))
      (ss-draw-all) )))

(defun ss-move-up ()
  (interactive)
  (if (> ss-cur-row 1)
      (ss-move-cur-cell 0 -1) nil ))
    
(defun ss-move-down ()
  (interactive)

  (if (< ss-cur-row (- ss-max-row 1))
      (ss-move-cur-cell 0 1 )
    (progn
      (setq ss-max-row (+ 1 ss-max-row))
      (setq ss-cur-row (+ 1 ss-cur-row))
      (ss-draw-all) )))

(defun ss-move-cur-cell (x y) (interactive)
  (let* ((new-row (+ ss-cur-row y))
	 (new-col (+ ss-cur-col x))
	 (m (avl-tree-member ss-data (+ ss-cur-row (* ss-max-row (+ ss-cur-col 1)))))
	 (n (avl-tree-member ss-data (+ new-row (* ss-max-row (+ new-col 1))))) 
	 (ot (if m (elt m 1) ""))
	 (nt (if n (elt n 1) "")))
    (progn
    (ss-draw-cell ss-cur-col ss-cur-row  (ss-pad-right ot ss-cur-col))
    (setq ss-cur-row new-row)
    (setq ss-cur-col new-col)
    (ss-draw-cell ss-cur-col ss-cur-row  (propertize  (ss-pad-right nt ss-cur-col) 'font-lock-face '(:inverse-video t))) )))
  
  

(defun ss-edit-cell ()
  "edit the selected cell"
  (interactive)
  (let*  ( (m (avl-tree-member ss-data (+ ss-cur-row (* ss-max-row (+ ss-cur-col 1)))))
	   (ot (if m (elt m 1) ""))
	   (nt (read-string "Cell Value:" ot )) )
    (if (= (string-to-char "=") (elt nt 0)) ; new value starts with  =
	(nil) ;; add funciton
      (if m (aset m 1 nt)
	(progn  ;;else
	  (setq m (vector (concat (ss-col-letter ss-cur-col) (int-to-string ss-cur-row)) nt))
	  (avl-tree-enter ss-data m)  )))
    (ss-draw-cell ss-cur-col ss-cur-row (propertize (elt m 1) 'font-lock-face '(:inverse-video t))) ))

(defun ss-set-function (fun)
  "Sets the function for current cell to value"
  
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
;; 	     (line (thing-at-point 'line)))
;; 	 (dolist (cell (split-string line ","))
;; 	   (if (and (char-equal "\"" (substring cell 0 0))
;; 		    (char-equal "\"" (substring cell -1 -1)))
;; 	       (setq cell (substring cell 1 -1)) nil )
;; 	   (setq row (vconcat row (list  cell))) )
;; 	 (setq ss-mode-data (vconcat ss-mode-data row))
;; 	 (with-current-buffer (ses-goto-print (x+1) 0))
;; 	 (next-line) 
;; 	 ))) )



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
  (setq cursor-type nil)
  (setq ss-cur-col 0)
  (setq ss-max-col 3)
  (setq ss-col-widths (make-vector ss-max-col 7))
  (setq ss-cur-row 1)
  (setq ss-max-row 3)
  (setq ss-row-padding 4)
  (setq ss-data (avl-tree-create 'ss-avl-cmp))

  (ss-draw-all))


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

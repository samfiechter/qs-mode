;;; ss-mode -- Spreadsheet Mode -- Tabular interface to Calc
;; Copyright (C) 2014 -- Use at 'yer own risk  -- NO WARRANTY!
;; Author: sam fiechter sam.fiechter(at)gmail
;; Version: 0.000000000000001
;; Created: 2014-03-24
;; Keywords: calc, tabular 


;;;Code

;; ;;;;;;;;;;;;; variables ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defface spread-cell-face
  '((((class grayscale) (background light)) (:foreground "LightGray" :bold t))
    (((class grayscale) (background dark)) (:foreground "DimGray" :bold t))
    (((class color) (background light)) (:foreground "Purple"))
    (((class color) (background dark)) (:foreground "Blue" :bold t))
    (t (:bold t)))
  "Spread-mode face used to highlight cells"
  :group 'spread-mode)

(defvar ss-mode-empty-name "*Sams Spreadsheet Mode*")
(defvar ss-mode-column-widths (list ))
(defvar ss-mode-cur-col 0)
(defvar ss-mode-cur-row 0)

;; ;;;;;;;;;;;;; keymaps ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defvar ss-mode-map  (make-sparse-keymap 'ss-mode-map))


(define-key ss-mode-map "q"                'ss-close)

 (define-key ss-mode-map [left]        'backward-sexp)
 (define-key ss-mode-map [right]        'forward-sexp)
 (define-key ss-mode-map [up]                'previous-line)
 (define-key ss-mode-map [down]        'next-line)
 (define-key ss-mode-map (kbd "RET")        'ss-edit-cell)



(defvar ss-null-map
  (make-sparse-keymap 'ss-null-map))
(define-key ss-null-map "n"                'ss-start-game)

;; ;;;;;;;;;;;;; functions ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; SAM's SES MACROS

;;;###autoload
(define-derived-mode ss-mode tabulated-list-mode ss-mode-empty-name
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
  (setq tabulated-list-format [("" 4 t)]) ;; 0th row is for numbers
  (let ((a 0) (w nil))
    (dotimes (a 3 )
      (setq w (elt ss-mode-column-widths a))     
      (setq tabulated-list-format (vconcat tabulated-list-format (list (list (char-to-string (+ ?A a)) 
									 12 t)))) ))
  (setq tabulated-list-padding 0)
  (pop-to-buffer ss-mode-empty-name nil)
  (setq cursor-type nil)
  (setq tabulated-list-entries (list (list "1" [ "1" "1" "2" "3"] ) 
				     (list "2" [ "33333333333333333333333332" "4" "5" "6" ] )))
  (tabulated-list-init-header) 
  (tabulated-list-print t))

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

  (ss-mode)
)



(defun ss-edit-cell ()
(interactive)
(let row (elt tabulated-list-entry 1)
     (setq row tabulated-list-get-id (read-string "Cell Value:" (aref row tabulated-list-get-id)))
     (tabulated-list-print t) ))


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

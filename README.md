# SudokuHelper
A Sudoku Helper in Excel via VBA
Please note that It's just a helper, not a solver!

1. Start a new Workbook
2. Insert the codes into Sheet1(codec)
3. Run `Initialize`

A. Clear the puzzle
B. Label figures as Fixed(Level:5) or Sured(Level:4)
C. For any cell with Level < 4, if the cell contains only 1 figure, Label it as Confirmed(Level:3)
D. For any cell with Level < 4, clear the cell
E. For any cell with Level < 4, guess what figures are ok

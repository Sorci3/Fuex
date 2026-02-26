from tkinterdnd2 import TkinterDnD
from source.ui import ExcelMergerUI

def main():
    root = TkinterDnD.Tk()
    
    app = ExcelMergerUI(root)
    
    root.mainloop()

if __name__ == "__main__":
    main()
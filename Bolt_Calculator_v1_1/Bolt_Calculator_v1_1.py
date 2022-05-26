from tkinter import Tk
from main import MainWindow

def main() -> None:
    """
    Starts Bolt Calculator
    ...

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
    master = Tk()
    app = MainWindow(master)
    app.mainloop()


if __name__ == "__main__":
    main()

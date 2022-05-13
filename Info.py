from pathlib import Path
from tkinter import Frame, IntVar, Toplevel, ttk
from typing import Tuple


class Info(Frame):
    """
    Class defining Information Window, showing calculation results.
    """

    def __init__(self, master=None, cwd: Path = None, bolt=None):
        Frame.__init__(self, master)

        self.master: Toplevel = master

        self.master.title("Calculation result")
        self.master.iconbitmap(cwd.joinpath("docs", "b_c_logo.ico"))
        self.master.resizable(False, False)

        self._hF = Frame(master)
        self._hF.pack(fill="both", expand=1)

        self._lF = Frame(master)
        self._lF.pack(fill="both", expand=1)

        self._cbF = Frame(master)
        self._cbF.pack(fill="both", expand=1)

        self._bF = Frame(master)
        self._bF.pack(fill="both", expand=1)

        ttk.Style().configure("TLabel", font=("Calibri", 12))

        self._lb_header = ttk.Label(
            self._hF, text="Calculation results for bolt dimensioning:", font=("Calibri", 12, "bold"))
        self._lb_header.pack(fill="both", pady=5, padx=5,)

        self._lb_std = ttk.Label(self._lF, text="Standard:")
        self._lb_std.grid(row=0, column=0, pady=2, padx=5, stick="nsew")

        self._lb_std_val = ttk.Label(
            self._lF, text=self._get_standard(bolt))
        self._lb_std_val.grid(row=0, column=1, pady=2, padx=5, stick="nsew")

        self._lb_diameter = ttk.Label(self._lF, text="Diameter:")
        self._lb_diameter.grid(row=1, column=0, pady=2, padx=5, stick="nsew")

        self._lb_diameter_val = ttk.Label(
            self._lF, text=str(bolt.diameter) + " mm")
        self._lb_diameter_val.grid(
            row=1, column=1, pady=2, padx=5, stick="nsew")

        self._lb_length = ttk.Label(self._lF, text="Length:")
        self._lb_length.grid(row=2, column=0, pady=2, padx=5, stick="nsew")

        self._lb_length_val = ttk.Label(
            self._lF, text=str(bolt.length) + " mm")
        self._lb_length_val.grid(
            row=2, column=1, pady=2, padx=5, stick="nsew")

        self._lb_mat = ttk.Label(self._lF, text="Material:")
        self._lb_mat.grid(row=3, column=0, pady=2, padx=5, stick="nsew")

        self._lb_mat_val = ttk.Label(self._lF, text=bolt.material.name)
        self._lb_mat_val.grid(row=3, column=1, pady=2, padx=5, stick="nsew")

        self._CAD = IntVar(value=0)
        self._report = IntVar(value=0)

        self._check_CAD = ttk.Checkbutton(
            self._cbF, text="Create CAD model", variable=self._CAD)
        self._check_CAD.grid(row=0, column=0, pady=10, padx=5)

        self._check_rep = ttk.Checkbutton(
            self._cbF, text="Create report", variable=self._report)
        self._check_rep.grid(row=0, column=1, pady=10, padx=5)

        self._btn_pro = ttk.Button(
            self._bF, text="Proceed", command=self._proceed)
        self._btn_pro.pack(anchor="center", pady=5)

    def _get_standard(self, bolt) -> str:
        """
        Set label displaying selected bolt standard.
        ...

        Parameters
        ----------
        bolt: Bolt
            calculated bolt

        Returns
        standard: str
            string representing the selected bolt
        """
        if "ISO2340" in str(type(bolt)):
            return (f"ISO 2340 {bolt.form}")
        elif "ISO2341" in str(type(bolt)):
            return (f"ISO 2341 {bolt.form}")
        else:
            return "DIN 1445"

    def _proceed(self) -> None:
        """
        Closing Information window to proceed with programm execution.
        ...

        Parameters
        ----------
        None

        Returns
        -------
        None
        """
        self.master.destroy()

    def to_main(self) -> Tuple[IntVar]:
        """
        Returning value of checkboxes to decide wheather CAD model and/ or 
        report are created
        ...

        Parameters
        ----------
        None

        Returns
        -------
        tuple: Tuple(IntVar)
            tuple containing IntVars showing if CAD model and report are to
            be created or not.
        """
        self.wait_window()
        return (self._CAD.get(), self._report.get())

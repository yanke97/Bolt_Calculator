import getpass
from datetime import date
from os import startfile
from pathlib import Path
from typing import Callable, Dict

from fpdf import FPDF

from Bolt import Bolt
from Calculation import Calculation


class PathError(Exception):
    """
    Custom Exception. Raised when provided path does not exist
    """

    def __init__(self, path: Path) -> None:
        self._path = path
        self._message = f"The given path:\n\n {self._path}\n\ndoes not exist. Please check report output path in settings."

        super().__init__(self._message)


class Report(FPDF):
    """
    CLass definining a calculation report

    ...

    Attributes
    ----------
    _bolt_info: Dict[str, str]
        Dictionary containing necessary information about the calculated bolt
        as strings
    _mat_info: Dict[str, str]
        Dictionary containing necessary information about the used material
        as strings
    _calc_info: Dict[str, str]
        Dictionary containing necessary information about the perfomred 
        calculation as strings
    _path: str
        path to where the report is to be safed
    _function: Callable
        function processing status information


    Methods
    -------
    header() !! not be called explicitly !!
    footer() !! not be called explicitly !!
    create()

    """

    def __init__(self, cwd: Path, bolt: Bolt, calc: Calculation, path: str, function: Callable):
        self._cwd: Path = cwd
        self._bolt_info: Dict[str, str] = bolt.report_dict()
        self._mat_info: Dict[str, str] = bolt.material.report_dict()
        self._calc_info: Dict[str, str] = calc.report_dict()
        self._name: str = self._bolt_info["Name:"]
        self._path: Path = Path(path)
        self._function: Callable = function

        super().__init__(orientation="P", unit="mm", format="A4")

        self.set_left_margin(20)
        self.add_page()
        self.set_margins(20, 15, 15)
        self.line(5, 297/2, 10, 297/2)
        self._create()

    def header(self) -> None:
        """
        Defining the report header
        ...

        Parameters
        ----------
        None

        Returns
        -------
        None
        """

        today = date.today().strftime("%d.%m.%Y")

        self.set_font('Helvetica', '', 12)
        self.image(str(self._cwd.joinpath("docs", "b_c_logo.png")), w=10, h=10)
        self.set_xy(30, 10)
        self.multi_cell(30, 5, "Bolt\nCalculator", 0)
        self.set_xy(80, 10)
        self.set_font('Helvetica', 'B', 15)
        self.cell(50, 10, 'Calculation Report', 0, 0, 'C')
        self.set_font('Helvetica', "", 12)
        self.multi_cell(
            60, 5, f"Date: {today}\nCreator: {getpass.getuser()}", 0, "R")
        self.line(20, 23, 195, 23)
        self.ln(20)

    def footer(self) -> None:
        """
        Defining the report footer
        ...

        Parameters
        ----------
        None

        Returns
        -------
        None
        """
        self.set_y(-15)
        self.set_font('Helvetica', '', 5)
        self.multi_cell(
            60, 2, "Created using Bolt Calculator V2. \nComputations performed accoridng to Roloff/ Matek 23rd edt.", 0)
        self.set_y(-15)
        self.set_font('Helvetica', '', 10)
        self.cell(0, 10, f"{self.page_no()}/1", 0, 0, 'R')

    def _create(self) -> None:
        """
        Create report
        ...

        Paratmetes
        ----------
        None

        Returns
        -------
        None
        """
        # Write General information
        self.set_font("Helvetica", "U", 12)
        self.write(5, "General Information:")
        self.set_font("Helvetica", "", 12)
        self.ln(5)
        for item in self._bolt_info.items():
            self.ln(5)
            self.set_x(25)
            self.cell(70, 5, f"{item[0]}")
            self.cell(40, 5, f"{item[1]}")
        self.ln(15)

        # Wirte material information
        self.set_font("Helvetica", "U", 12)
        self.write(5, "Material Information:")
        self.set_font("Helvetica", "", 12)
        self.ln(5)
        for item in self._mat_info.items():
            self.ln(5)
            self.set_x(25)
            self.cell(70, 5, f"{item[0]}")
            self.cell(40, 5, f"{item[1]}")
        self.ln(15)

        # Write Computational information
        self.set_font("Helvetica", "U", 12)
        self.write(5, "Computational Information:")
        self.set_font("Helvetica", "", 12)
        self.ln(5)

        for i, item in enumerate(self._calc_info.items()):
            self.ln(5)
            self.set_x(25)
            self.cell(70, 5, f"{item[0]}")
            self.cell(40, 5, f"{item[1]}")
            if i == 5 or i == 10 or i == 13 or i == 15:
                self.ln(5)

        self._save_report()

    def _save_report(self) -> None:
        """
        Save and open report.

        If a report with same name is opened a counter is added to the name
        ...

        Parameters
        ---------
        None

        Returns
        ------
        None

        Raises
        ------
        PathError
            if given path does not exist
        """
        name = "Calculaion_Report_{}{}.pdf"
        try:
            self.output(
                str(self._path/name.format(self._bolt_info["Name:"], "")))
            startfile(
                str(self._path/name.format(self._bolt_info["Name:"], "")))
        except PermissionError:
            i = 0
            while True:
                i += 1
                path = self._path/name.format(self._bolt_info["Name:"], i)
                if not path.exists():
                    self.output(str(path))
                startfile(
                    str(self._path/name.format(self._bolt_info["Name:"], "")))
                return
        except OSError:
            self._function("PathError occured.")
            raise PathError(self._path) from None

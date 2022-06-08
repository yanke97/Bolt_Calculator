from dataclasses import dataclass, field
from os import path
import win32com.client as win32
# Import class for objects returned from client.Dispatch function
from win32com.client import CDispatch
# Import the com_error, an error raised when a COM exception occures
from pywintypes import com_error


class ExcelFileNotFoundError(com_error):
    """
    Custom FileNotFoundError. Raised when provided path to Excel file is 
    erroneous.
    """

    def __init__(self, file_path: str) -> None:
        self._path = file_path
        self._message = f"Error when trying to open Excel file: \
                        \n {self._path}. \nPlease check the filepath in Settings."

        super().__init__(self._message)


@dataclass
class Connector:
    """
    class handling Excel connection

    ...
    Attributes
    ----------
    _worbook_path: str
        path to the workbook containing bolt dimensions
    _excel_application: CDispatch
        Excel application in which the workbook is opened
    _workbook_name: str
        name of the workbook containing the bolt dimensions
    _worbook: CDispatch
        workbook containing bolt dimensions
    _worksheet: CDispatch
        worksheet containing bolt dimensions
    _was_running: bool 
        variable storing information about whether Excel was running or not.
        Default is False
    _was_opened: bool
        variable storing information about whether the workbook was opened or not.
        Default is False

    Methods
    -------
    connect_to_excel()
    close_excel()

    """

    _worbook_path: str
    _excel_application: CDispatch = field(init=False)
    _workbook_name: str = field(init=False)
    _worbook: CDispatch = field(init=False)
    _worksheet: CDispatch = field(init=False)
    _was_running: bool = False
    _was_opened: bool = False

    def __post_init__(self):
        self._workbook_name = path.basename(self._worbook_path)

    def connect_to_excel(self) -> CDispatch:
        """
        Connect to Excel and return the worksheet containing bolt dimensions

        ...

        Parameters
        ----------
        path: str
            path to the excel worbook containing the bolt dimensions

        Returns
        -------
        ws: object
            worksheet containing bolt dimensions

        Raises
        ------
        ExcelFileNotFoundError
            raised when excel file can not be found under the given path 

        """
        try:
            self._excel_application = win32.GetActiveObject(
                "Excel.Application")
            self._was_running = True
        except com_error:
            self._excel_application = win32.gencache.EnsureDispatch(
                "Excel.Application")

        if self._was_running is False:
            try:
                self._workbook = self._excel_application.Workbooks.Add(
                    self._worbook_path)
            except com_error:
                self._excel_application.Quit()
                raise ExcelFileNotFoundError(self._worbook_path) from None
        else:
            for w in self._excel_application.Workbooks:
                if w.Name == self._workbook_name:
                    self._excel_application.ActiveWorkbook = w
                    self._workbook = self._excel_application.ActiveWorkbook
                    self._was_opened = True

            if self._was_opened is False:
                try:
                    self._workbook = self._excel_application.Workbooks.Add(
                        self._worbook_path)
                except com_error:
                    raise ExcelFileNotFoundError(
                        self._worbook_path) from None

        self._worksheet = self._workbook.ActiveSheet
        return self._worksheet

    def close_excel(self) -> None:
        """
        Close running Excel application

        ...

        Parameters
        ----------
        None

        Returns
        -------
        None
        """

        if self._was_opened is True:
            self._workbook.Close(False)
        else:
            self._workbook.Close(False)
            self._excel_application.Quit()

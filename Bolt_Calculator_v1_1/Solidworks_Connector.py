import os
from typing import Dict

import pythoncom
import win32com.client as win32
# Import class for objects returned from client.Dispatch function
from win32com.client import CDispatch

from Bolt import Bolt


class CADFileNotFoundError(FileNotFoundError):
    """
    Custom FileNotFoundError. Raised when provided path to CAD model is 
    erroneous.
    """

    def __init__(self, path: str) -> None:
        self._path = path
        self._message = f"The CAD model with path:\n {self._path} \ncould not be found. Please check path in settings. \n\nModel creation stopped."
        super().__init__(self._message)


class CADFileError(Exception):
    """
    Custom Error. Raised when provided CAD file is not a part.
    """

    def __init__(self, path: str) -> None:
        self._path = path
        self._message = f"The provided CAD file:\n {os.path.basename(self._path)}\nis not of type part. \n\nModel creation stopped."
        super().__init__(self._message)


class CADFileWithSameNameAlreadyOpenError(Exception):
    """
    Custom Erro. Raised when a CAD file with the same name is already open.
    """

    def __init__(self, path: str) -> None:
        self._path = path
        self._message = f"A CAD file with name\n {os.path.basename(self._path)}\nis already opened. \n\nModel creation stopped."
        super().__init__(self._message)


class SldWorksConnectionError(Exception):
    """
    Custom Error. Raised when connection to SolidWorks could not 
    be established.
    """

    def __init__(self) -> None:
        self._message = "The connection to SolidWorks could not be established. Please make sure Solidworks is installed on your computer. \n\nModel creation stopped."
        super().__init__(self._message)


class VariableError(Exception):
    """
    Custom Error. Raised when the given SolidWorks part does not have any 
    global variables.
    """

    def __init__(self, path: str) -> None:
        self._model: str = os.path.basename(path)
        self._message = f"The provided CAD model \n {self._model}\ndoes not contain any global variables.\nPlease provide a suitable model. \n\nModel creation stopped."
        super().__init__(self._message)


class SaveError(Exception):
    """
    Custom Error. Raised when the given path for saving the 
    CAD model is erroneous or a part with the same name exists and is opened.
    """

    def __init__(self, path: str) -> None:
        self._path = path
        self._message = f"The provided path for svaing the model:\n {self._path}\ndoes not exist. Please make sure you use an exisiting path. Or a part with the same name exists and is opened. Please colse this part and try again. \n\nModel creation stopped."
        super().__init__(self._message)


class SldWorksConnector():
    """
    Class handling the connection to SolidWorks

    ...
    Attributes
    ----------
    _input_path: str 
        path to folder containing CAD template models
    _output_path: str
        path to folder to which the created CAD models are saved
    _bolt_dict: Dict[str, str]
        dictionary containing necessary information to be pased to the CAD model
    _sldwrks: CDispatch
        SolidWorks application
    _part: CDispatch
        CAD model which varaibles are to be set
    _eq_mngr: CDispatch
        equation manager of the CAD model

    Methods
    -------
    _connect_to_sldwrks()
    _set_variables()
    _get_global_variables()
    _save_part()
    main()

    """

    def __init__(self, input_path: str, output_path: str, bolt: Bolt):
        self._input_path: str = input_path
        self._output_path: str = output_path
        self._bolt_dict: Dict[str, int] = bolt.cad_dict()
        self._part_path = None
        self._sldwrks: CDispatch = None
        self._part: CDispatch = None
        self._eq_mngr: CDispatch = None

    def _connect_to_sldwrks(self, function) -> None:
        """
        Connect to SolidWorks and open the template model

        ...
        Parameters
        ----------
        function: callable
            function processing status information

        Returns
        -------
        None

        Raises
        ------
        CADFileNotFoundError
            when provided path to CAD model is erroneous

        CADFileError
            when provided CAD file is not a part

        CADFileWithSameNameAlreadyOpenError
            when a CAD file with the same name is already open

        """

        self._sldwrks = win32.Dispatch("SldWorks.Application")

        # Checking if connection was established, otherwise raise Error
        if self._sldwrks is None:
            raise SldWorksConnectionError from None

        self._part_path = self._input_path.replace(
            "\\", "/") + "/" + self._bolt_dict["standard"].replace(" ", "") + "_template.SLDPRT"

        # Checks if path to part does exists, returns true if it does
        if os.path.exists(self._part_path) is True:
            part_path = win32.VARIANT(pythoncom.VT_BSTR, self._part_path)
            # sets the type of the doc to be opened to "part"
            type = win32.VARIANT(pythoncom.VT_I4, 1)
            # sets the opening option to "silent"
            option = win32.VARIANT(pythoncom.VT_I4, 1)
            # Variable that contains the errors that may occure during opening the file
            # A int identifying an error type must be given, but all errors of diffent types are catched as well
            error = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 2)
            # Variable that contains the warnings that may be shown during opening the file
            # A int identifying a warning type must be given, but all warnings of diffent types are catched as well
            warning = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128)

            self._part = self._sldwrks.OpenDoc6(
                part_path, type, option, "", error, warning)

            # Checking if error occured while opening the part
            if error.value != 0:
                if error.value == 2:
                    raise CADFileNotFoundError(part_path) from None
                elif error.value == 2097152:
                    raise CADFileError(part_path) from None
                elif error.value == 65536:
                    raise CADFileWithSameNameAlreadyOpenError(
                        part_path) from None
                else:
                    quit()

            if warning.value != 0:
                if warning.value == 128:
                    function("CAD-file is already opened.")

            self._sldwrks.Visible = True
        else:
            raise CADFileNotFoundError(self._part_path) from None

    def _set_variables(self) -> None:
        """
        Set gloabal variables in SolidWorks model to values of the bolt object.

        ...
        Parameters
        ----------
        None

        Returns
        -------
        None
        """

        global_variables = self._get_global_variables()

        # When iterationg over dicts keys are accessed by default
        # To iterate ocer items the items must be accessed explicitly by using .items()
        for key in global_variables:
            self._eq_mngr.Equation(global_variables[key],
                                   f"\"{key}\"={self._bolt_dict[key]}")

        Nothing = win32.VARIANT(pythoncom.VT_DISPATCH, None)

        try:
            if self._bolt_dict["hole"] is True:
                self._part.Extension.SelectByID2(
                    "Schnitt-Linear austragen1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
                self._part.EditUnsuppressDependent2
        except KeyError:
            pass

        self._part.EditRebuild3

    def _get_global_variables(self) -> Dict[str, int]:
        """
        Retrieve all global variables from the CAD-Model and store them in a dictionary.
        Names are used as keys, values represent variable index in list of global variables

        Parameters
        ----------
        None

        Returns
        -------
        global_variables: Dict[str, int]
            dictionary containing all global variables of the model and thei index

        Raises
        ------
        VariableError
            raised when CAD model does not have any global variables

        """
        global_variables = {}

        self._eq_mngr = self._part.GetEquationMgr

        for i in range(self._eq_mngr.GetCount):
            if self._eq_mngr.GlobalVariable(i) is True:
                global_variables[self._eq_mngr.Equation(i).split("\"")[1]] = i

        if len(global_variables.keys()) == 0:
            self._sldwrks.CloseDoc(os.path.basename(self._part_path))
            raise VariableError(self._input_path) from None
        else:
            return global_variables

    def _save_part(self) -> None:
        """
        Save part

        ...
        Parameters
        ----------
        None

        Returns
        -------
        None

        Raises
        ------
        SaveError
            when the given path for saving the CAD model is erroneous 
            or a part with the same name exists and is opened.

        """
        output_path = self._output_path.replace(
            "\\", "/") + "/" + self._bolt_dict["name"] + ".SLDPRT"

        arg = win32.VARIANT(pythoncom.VT_BSTR, output_path)

        status = self._part.SaveAs3(arg, 0, 0)

        if status != 0:
            if status == 1:
                self._sldwrks.CloseDoc(os.path.basename(self._part_path))
                raise SaveError(output_path) from None

    def main(self, function) -> None:
        """
        Handle all the actions to create and save the CAD model 
        of the calculated bolt.

        ...

        Parameters
        ----------
        function: callable
            function processing status information

        Returns
        -------
        None

        """
        function("Connecting to SolidWorks...")
        self._connect_to_sldwrks(function)
        function("Connection to SolidWorks established")
        self._set_variables()
        function("Dimensions of CAD model set")
        self._save_part()
        function("Model saved")
        function("Connection to SolidWorks closed")

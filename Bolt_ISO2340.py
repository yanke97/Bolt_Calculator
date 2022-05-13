from __future__ import annotations
from dataclasses import dataclass, field
from typing import Dict, Any
from win32com.client import CDispatch
from Bolt import Bolt


@dataclass
class BoltISO2340(Bolt):
    """
    Class for bolts according to DIN 14453.
    Inherits from abstractbaseclass Bolt

    ...

    Attributes
    ----------
    name: str
        the name of the bolt
    material: Material
        bolt material
    diameter: int init=False
        bolt diameter [mm]
    length: int init=False
        bolt length [mm]
    standard: str init=False
        string representing the standard in accordance to which the bolt
        dimensions will be set
    _chamfer_height_30: int init=False
        heigth of the 30° chamfer of the bolt [mm]
    _hole_diameter: int init=False
        diameter of the pin hole [mm]
    _hole_dist: int init=False
        distance of the pin hole from the end surfaces [mm]
    form: str
        string defining wether the bolt has a pin hole (fomr B) or not (form A)

    Methods
    -------
    set_length()
    set_dimensions()
    cad_dict()
    report_dict()


    """
    form: str
    _chamfer_height_30: float = field(init=False)
    _hole_diameter: float = field(init=False)
    _hole_dist: float = field(init=False)

    def __post_init__(self) -> None:
        self._standard = "ISO 2340"

    def set_dimensions(self, ws: CDispatch) -> None:
        """
        Set bolt dimensions

        ...

        Parameters
        ----------
        ws: CDispatch
            Excel Worksheet object that contains the bolt dimensions

        Returns
        -------
        None

        """

        self._chamfer_height_30 = ws.Range("B10").Value
        self._hole_diameter = ws.Range("C10").Value
        self._hole_dist = ws.Range("D10").Value

    def cad_dict(self) -> Dict(str, Any):
        cad_dict = super().cad_dict()

        cad_dict["chamfer_height_30"] = self._chamfer_height_30
        cad_dict["hole_diameter"] = self._hole_diameter
        cad_dict["hole_distance"] = self._hole_dist

        if self.form == "A":
            cad_dict["hole"] = False
        else:
            cad_dict["hole"] = True

        return cad_dict

    def report_dict(self) -> Dict[str, str]:
        """
        Create a dict with all the information to be represented in a
        calculation report as strings.

        !!Only to be called after calculation is completed.!!

        Parameters
        ----------
        None

        Returns
        -------
        report_dict: Dict[str,str]
            Dictionary containing the information to be represented
        """

        report_dict = super().report_dict()

        report_dict.setdefault(
            "Form:", f"{self.form}")

        return report_dict

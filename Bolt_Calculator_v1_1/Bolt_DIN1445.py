from __future__ import annotations
from dataclasses import dataclass, field
from typing import Dict, Any
from win32com.client import CDispatch
from Bolt import Bolt


@dataclass
class BoltDIN1445(Bolt):
    """
    Class for bolts according to DIN 14453.
    Inherits from abstractbaseclass Bolt

    ...

    Attributes
    ----------
    _name: str
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
    _head_diameter: float init=False
        diameter of bolt head [mm]
    _head_thickness: float init=False
        thickness of bolt head [mm]
    _wrench_size: float init=False
        size of wrench applicable [mm]
    _radius: float init=False
        radius between bolt and bolt head [mm]
    _chamfer_height: float init=False
        height of the chamfer at the bolt head [mm]
    _thread_diameter: float init=False
        diameter of the thread at the end of the bolt [mm]
    _thread_length: float init=False
        length of the thread [mm]
    _undercut_diameter:float init=False
        diameter of the undercut [mm]
    _undercut_length: float init=False
        length of the undercut [mm]
    _undercut_radius: float init=False
        radius applied to the undercut [mm]

    Methods
    -------
    _set_length()
    _set_dimensions()
    cad_dict()
    report_dict()
    """

    _head_diameter: float = field(init=False)
    _head_thickness: float = field(init=False)
    _wrench_size: float = field(init=False)
    _radius: float = field(init=False)
    _chamfer_height: float = field(init=False)
    _thread_diameter: float = field(init=False)
    _thread_length: float = field(init=False)
    _undercut_diameter: float = field(init=False)
    _undercut_length: float = field(init=False)
    _undercut_radius: float = field(init=False)

    def __post_init__(self) -> None:
        self._standard = "DIN 1445"

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

        self._chamfer_height = ws.Range("B10").Value
        self._head_diameter = ws.Range("E10").Value
        self._head_thickness = ws.Range("F10").Value
        self._radius = ws.Range("G10").Value
        self._chamfer_height = ws.Range("H10").Value
        self._wrench_size = ws.Range("I10").Value
        self._thread_length = ws.Range("J10").Value
        self._thread_diameter = ws.Range("K10").Value
        self._undercut_diameter = ws.Range("L10").Value
        self._undercut_length = ws.Range("M10").Value
        self._undercut_radius = ws.Range("N10").Value

    def cad_dict(self) -> Dict(str, Any):
        cad_dict = super().cad_dict()

        cad_dict["head_diameter"] = self._head_diameter
        cad_dict["head_thickness"] = self._head_thickness
        cad_dict["wrench_size"] = self._wrench_size
        cad_dict["radius"] = self._radius
        cad_dict["chamfer_height"] = self._chamfer_height
        cad_dict["thread_diameter"] = self._thread_diameter
        cad_dict["thread_length"] = self._thread_length
        cad_dict["undercut_diameter"] = self._undercut_diameter
        cad_dict["undercut_length"] = self._undercut_length
        cad_dict["undercut_radius"] = self._undercut_radius

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
            "Thread:", f"M{round(self._thread_diameter,0)}x{self._thread_length}")

        return report_dict

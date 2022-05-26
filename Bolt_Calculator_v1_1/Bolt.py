from __future__ import annotations
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Dict, Any
from win32com.client import CDispatch
from Material import Material


@dataclass
class Bolt(ABC):
    """
    Abstractbaseclass for Bolt types. Not to be instantiated.

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
    standard: str
        string representing the standard in accordance to which the bolt
        dimensions will be set

    Methods
    -------
    set_dimensions() (abstractbasemethod)
    cad_dict() (abstractmethod)
    report_dict() (abstractmethod)
    set_length()

    """

    _name: str
    material: Material
    diameter: float = field(init=False)
    length: float = field(init=False)
    _standard: str = field(init=False)

    @abstractmethod
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

    @abstractmethod
    def cad_dict(self) -> Dict(str, Any):
        """
        Create a dict with all the information to create the CAD-model
        !!Only to be called after calculation is completed.!!

        Parameters
        ----------
        None

        Returns
        -------
        report_dict: Dict[str,str]
            Dictionary containing the information to create the CAD-model
        """
        cad_dict = {"name": self._name,
                    "standard": self._standard,
                    "diameter": self.diameter,
                    "length": self.length}
        return cad_dict

    @abstractmethod
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

        report_dict = {"Name:": self._name,
                       "Standard:": self._standard,
                       "Material:": self.material.name,
                       "Diameter [mm]:": str(self.diameter),
                       "Length [mm]:": str(self.length)}

        return report_dict

    def set_length(self, ws: CDispatch, n: int, t_r: float, t_f: float, i: int = 0) -> None:
        """
        Set the bolt length to a length provided by standard DIN 1445

        ...

        Parameters
        ----------
        ws: CDispatch
            Excel Worksheet object that contains the bolt dimensions
        i: int
            iterator, default is 0
        n: int
            number of sheers of the bolt joint
        t_r: float
            rod thickness [mm]
        t_f: float
            thickness of one part of the fork [mm] 

        Returns
        -------
        None

        """

        i += 1
        ws.Range("B3").Value = self._standard
        ws.Range("C3").Value = t_r
        ws.Range("D3").Value = t_f
        ws.Range("E3").Value = n
        length = ws.Range("O10").Value
        if length == 0:
            length = t_r + (n * t_f) + i
            ws.Range("A7").Value = length
            return self.set_length(ws, n, t_r, t_f, i)
        else:
            self.length = length

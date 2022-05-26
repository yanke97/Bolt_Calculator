from __future__ import annotations
from dataclasses import dataclass, asdict
from typing import Any, List, Dict
from csv import DictReader, Sniffer, DictWriter
from enum import Enum, auto


class MaterialFileNotFoundError(Exception):
    """
    Custom FileNotFoundError. Raised when the provided .mat-file can not be 
    found.
    """

    def __init__(self, file: str) -> None:
        self._file = file
        self._message = f"Error when trying to load Material information:\n\n {self._file} \
                \n\nFile not found. Please check filepath in settings"

        super().__init__(self._message)


class MaterialFileError(Exception):
    """
    Custom Error. Is raised when the file provided is not a .mat-file
    """

    def __init__(self, file: str) -> None:
        self._file = file
        self._message = "The given file is not of a supported material file type."

        super().__init__(self._message)


class MaterialTypeError(Exception):
    """
    Custom Error. Is raised when a material is not of a supported material type
    """

    def __init__(self, name: str, mattype: str) -> None:
        self._name = name
        self._mattype = mattype
        self._message = f"Material {self._name} has undefined material type:\
            \n    {self._mattype}"

        super().__init__(self._message)


class MaterialType(Enum):
    """
    Class defining the available material types.
    """

    STRUCTURAL_STEEL = auto()
    HEAT_TREATABLE_STEEL = auto()
    NITRIDING_STEEL = auto()


@dataclass(frozen=True)
class Material:
    """
    Dataclass used to represent materials.
    Attributes can only be set on instanciation.
    Class is not to be inherited.

    ...

    Attributes
    ----------
    name: str
        the name of the material
    matnr: str
        the materialnumber according to EN 10027-2:1992-09
    density: float
        material density in [kg/m³]
    tensilestrength: float
        the materials tensile strength in [N/mm²]
    yieldstress: float
        the materials yield stress in [N/mm²]
    youngsmodulus: float
        the materials youngs modulus in [N/mm²]
    type: str
        the type of material, e.g.: structural steel

    Methods
    -------
    to_string()
        return the mateial name and yield stress

    to_str_nr()
        return the material number and yield stress
    """

    __slots__ = ["name", "matnr", "density", "tensilestrength",
                 "yieldstress", "youngsmodulus", "type"]
    name: str
    matnr: str
    density: float
    tensilestrength: float
    yieldstress: float
    youngsmodulus: float
    type: MaterialType

    def __post_init__(self):
        #type = MaterialType[str(self.type)]
        
        if isinstance(self.type, MaterialType) is False:
            raise MaterialTypeError(self.name, self.type)

    # ---Methods---
    def to_string(self) -> str:
        """
        Return Name and yieldstress in a string

        ...

        Parameters
        ----------
        None

        Returns
        -------
        str
            a string containing the material name and yield stress
        """
        return (f"{self.name} ({round(self.yieldstress,0)} N/mm²)")

    def to_string_nr(self) -> str:
        """
        Return Materialnumber and yieldstress in a string

        ...

        Parameters
        ----------
        None

        Returns
        -------
        str
            a string containing the materialnumber and yield stress
        """
        return (f"{self.matnr} ({round(self.yieldstress,0)} N/mm²)")

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

        report_dict = {"Name:": self.name,
                       "Materialnumber:": self.matnr,
                       "Denisty [kg/m³]:": self.density,
                       "Tensile strength [N/mm²]:": str(self.tensilestrength),
                       "Yield Stress [N/mm²]:": str(self.yieldstress),
                       "Youngs Modulus [N/mm²]:": str(self.youngsmodulus),
                       "Type:": self.type.name.replace("_", " ").lower()}

        return report_dict

    @staticmethod
    def mat_to_file(path: str, mat: Material) -> None:
        """
        Safe the given Materialobject to the given .mat-file

        ...

        Parameters
        ----------
        path: str
            path to where the .mat-file is saved
        mat: Material
            the material object to be saved

        Returns
        -------
        None

        Raises
        ------
        FileNotFoundError
            if the given .mat-file can not be found/ opened
        """

        try:
            # r+ opens file for reading and writing
            file = open(path, "r+", newline="", encoding="utf-8")
        except FileNotFoundError:
            raise MaterialFileNotFoundError(path) from None

        else:
            dialect = Sniffer().sniff(file.read(1024))
            writer = DictWriter(file, Material.__slots__, dialect=dialect)
            writer.writerow(asdict(mat))
            file.close()

    @staticmethod
    def mats_from_file(path: str) -> List[Any]:
        """
        Read materials from given .mat-file and return list of materials

        ...

        Parameters
        ----------
        path: str
            path to where the file is saved

        Returns
        -------
        list: Material
            List containing the extracted Material objects

        Raises
        ------
        MaterialFileNotFoundError
            if the given file can not be found/opened
        MaterialFileError
            if the given file is not of supported .mat-file type
        MaterialTypeError
            if one or more materials are not of a supported material type
        """

        mats: List[Material] = []
        try:
            file = open(path, "r", newline="", encoding="utf-8")
        except FileNotFoundError:
            raise MaterialFileNotFoundError(path) from None
        else:
            dialect = Sniffer().sniff(file.read(1024))
            file.seek(0)
            header = Sniffer().has_header(file.read(1024))
            file.seek(0)

            if header is True:
                reader = DictReader(file, dialect=dialect)
                for r in reader:
                    try:
                        r["density"] = eval(r["density"])
                        r["tensilestrength"] = eval(r["tensilestrength"])
                        r["yieldstress"] = eval(r["yieldstress"])
                        r["youngsmodulus"] = eval(r["youngsmodulus"])
                        r["type"] = MaterialType[r["type"].split(".")[1]]
                        mats.append(Material(**r))
                    except KeyError:
                        raise MaterialTypeError(r["name"], r["type"]) from None

            else:
                raise MaterialFileError(path)

            file.close()
            return mats

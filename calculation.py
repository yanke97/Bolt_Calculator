from dataclasses import dataclass, field
from typing import List, Dict
import math
from Bolt import Bolt
from win32com.client import CDispatch
from Material import Material, MaterialType
from excel_connector import Connector


@dataclass
class Calculation:
    """
    Class handling all necessary calculations.

    ...

    Attributes
    ----------
    _bolt: Bolt
        bolt object which dimensions are calculated
    _l_type: str
        load type
    _case: str
        load case
    _F_r: float
        resulting force on the bolt [N]
    _t_r: float
        thickness of the rod [mm]
    _t_f: float
        thickness of the fork [mm]
    _n: int
        shear numer (1 or 2) [-]
    _K_A: float
        application factor [-]
    _S: float
        safty factor [-]
    _mat_rod: Material
        Material of the rod
    _mat_fork: Material
        Material of the fork
    _excel_path: str
        path to the excel file storing dimension information
    _l: List[float] init=False
        List containing the calculated correction factors according to the
        load type
    _stresses: List[float] init=False
        List containing all calculated maximum stresses as well as occuring
        stresses
    _M_b: float init=False
        bending moment acting upon the bolt
    _k: float init=False
        clamping factor according to clamping case
    _k_t: float init=False
        correcction Faktor for material thickness
    _yieldstress_min: float init=False
        minimal yieldstress of all materials in the bolt joint
    _connector: Connector init=False
        object of the Connector class used to comunicate with excel
    _ws: CDispatch init=False
        excel worksheet containing bolt dimensions

    Methods
    -------
    _calc_l() (staticmethod)
    _calc_m_b() (staticmethod)
    _calc_k_t()
    _calc_d_temp()
    _get_diameter()
    _check()
    main()

    """
    _bolt: Bolt
    _l_type: str
    _case: str
    _F_r: float
    _t_r: float
    _t_f: float
    _n: int
    _K_A: float
    _S: float
    _mat_rod: Material
    _mat_fork: Material
    _excel_path: str
    _l: List[float] = field(init=False)
    _stresses: List[float] = field(init=False)
    _M_b: float = field(init=False)
    _k: float = field(init=False)
    _k_t: float = field(init=False)
    _yieldstress_min: float = field(init=False)
    _connector: Connector = field(init=False)
    _ws: CDispatch = field(init=False)

    def __post_init__(self):
        self._l = self._calc_l(self._l_type)
        self._M_b, self._k = self._calc_m_b(
            self._case, self._F_r, self._t_r, self._t_f, self._n)

        if self._F_r < 0:
            self._F_r = self._F_r * (-1)

        self._k_t = 0
        self._stresses = []
        self._connector = Connector(self._excel_path)

        self._ws = self._connector.connect_to_excel()

    @staticmethod
    def _calc_l(l_type: str) -> List[float]:
        """
        Return list of loadfactors

        ...

        Parameters
        ----------
        lType: str
            string specifying the load type
            ("static", "alternating", "pulsating")

        Returns
        -------
        list
            list of floats representing the loadfactors for the
            respective stress types
            [normal stress, shear stress, pressure]
        """
        if l_type == "static":
            return [0.3, 0.2, 0.35]
        elif l_type == "pulsating":
            return [0.2, 0.15, 0.25]
        else:
            return [0.15, 0.1, 1]

    @staticmethod
    def _calc_m_b(case: str, F_r: float, t_r: float, t_f: float, n: int) -> List[float]:
        """
        Calculate the bending moment and clamping factor in dependency of
        the clamping case

        ...

        Parameters
        ----------
        case: str
            a string representing the clamping case
            ("Case 1", "Case 2", "Case 3")
        F_r: float
            the resulting force [N]
        t_r: float
            rod thickness [mm]
        t_f: float
            fork thickness [mm]
        n: int
            number of shears (1 or 2) [-]

        Returns
        -------
        list
            List of floats containing the bending moment
            and clamping factor
        """

        if case == "Case 1":
            return [(F_r * (t_r + n * t_f)) / 8, 1.6]
        elif case == "Case 2" and n == 2:
            return [(F_r * t_r) / 8, 1.1]
        elif case == "Case 2" and n == 1:
            return [F_r*t_r, 1.1]
        elif case == "Case 3" and n == 2:
            return [F_r*t_f, 1.1]
        else:
            return [F_r*t_f, 1.1]

    def _calc_k_t(self, d: float) -> None:
        """
        Calculate the material thickness correction factor in dependency of the
        material type.

        ...

        Parameters
        ----------
        d: float [mm]
            diameter of the bolt
        mat: Material
            bolt material

        Returns
        -------
        None

        """

        if self._bolt.material.type == MaterialType.HEAT_TREATABLE_STEEL:
            self._k_t = 1 - 0.41 * math.log10(d / 16)
            if self._k_t > 1:
                self._k_t = 1
            elif self._k_t < 0.60:
                self._k_t = 0.60
        elif self._bolt.material.type == MaterialType.STRUCTURAL_STEEL or \
                self._bolt.material.type == MaterialType.NITRIDING_STEEL:
            self._k_t = 1 - 0.23 * math.log10(d / 100)
            if self._k_t > 1:
                self._k_t = 1
            elif self._k_t < 0.89:
                self._k_t = 0.89

    def _calc_d_temp(self) -> float:
        """
        Calculate a temporary diameter for the bolt.

        ...

        Parameters
        ----------
        F_r:float
            resulting force on the bolt [N]
        l_type:str
            load type
        case:str
            clamping case
        t_r: float
            rod thickness [mm]
        t_f: float
            fork thickness [mm]
        n: int
            number of shears (1 or 2) [-]
        K_A:float
            application factor [-]
        S:float
            Safty factor [-]

        Returns
        -------
        d_temp:
            temporary diameter for the bolt [mm]

        """
        d_temp = self._k * math.sqrt((self._K_A * self._F_r * self._S) /
                                     (self._l[0] * self._bolt.material.yieldstress))

        return d_temp

    def _get_diameter(self, d_temp: float) -> float:
        """
        Set diameter in accordance to standard
        """

        self._ws.Range("A3").Value = d_temp
        d_std = self._ws.Range("A10").Value
        return d_std

    def _check(self, d: float, yieldstress_min: float) -> float:
        """
        Check if the bolt diameter is sufficient

        ...

        Parameters
        ----------
        d: float
            diameter to be checked [mm]
        yieldstress_min: float
            minimal yieldstress of all materials in the bolt joint [N/mm²]

        Returns
        -------
        d: float
            checked diameter [mm]
        """

        self._calc_k_t(d)

        # Calculating max. allowed stress
        sigma_max = self._k_t * self._bolt.material.yieldstress * self._l[0]
        self._stresses.insert(0, sigma_max)

        # Calculating bending stress:
        sigma = (self._K_A * self._M_b * 32) / (math.pi * d ** 3)
        self._stresses.insert(1, sigma)

        # Checking for safty
        if sigma_max / self._S <= sigma:
            d += 1
            d = self._get_diameter(d)
            return self._check(d, yieldstress_min)

        # Calculating max. alowed sheer stress:
        tau_max = self._k_t * self._bolt.material.yieldstress * self._l[1]
        self._stresses.insert(2, tau_max)

        # Calculating sheer stress:
        tau = 4 / 3 * (self._K_A * self._F_r) / (math.pi * (d ** 2 / 4) * 2)
        self._stresses.insert(3, tau)

        # Checking for safty
        if tau_max / self._S <= tau:
            d += 1
            d = self._get_diameter(d)
            return self._check(d, yieldstress_min)

        # Calculating max. allowed pressure:
        p_max = self._k_t * yieldstress_min * self._l[2]
        self._stresses.insert(4, p_max)

        # Calculating pressure in fork:
        p_f = (self._K_A * self._F_r) / (d * self._n * self._t_f)
        self._stresses.insert(5, p_f)

        # Calculating pressure in rod:
        p_r = (self._K_A * self._F_r) / (d * self._t_r)
        self._stresses.insert(6, p_r)

        # Checking for safty
        if p_max / self._S <= p_f and p_max * self._S <= p_r:
            d += 1
            d = self._get_diameter(d)
            return self._check(d, yieldstress_min)

        return d

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

        report_dict = {"Resulting Force [N]:": str(self._F_r),
                       "Safty Factor [-]:": str(self._S),
                       "Application Factor [-]:": str(self._K_A),
                       "Connection Type:": "double shear",
                       "Load Type:": self._l_type,
                       "Clamping Case:": "Case 1 (loose fit in rod and fork)",
                       "Load Factor Bending [-]:": str(self._l[0]),
                       "Load Factor Shear[-]:": str(self._l[1]),
                       "Load Factor Preassure [-]:": str(self._l[2]),
                       "Clamping Factor [-]:": str(self._k),
                       "Size Factor [-]:": str(self._k_t),
                       "Bending Moment [Nmm]:": str(self._M_b),
                       "allowed Bending Stress [N/mm²]:": str(self._stresses[0]),
                       "Bending Stress [N/mm²]:": str(round(self._stresses[1], 2)),
                       "allowed Shear Stress [N/mm²]:": str(self._stresses[2]),
                       "Shear Stress [N/mm²]:": str(round(self._stresses[3], 2)),
                       "allowed Preassure [N/mm²]:": str(self._stresses[4]),
                       "Preassure Fork [N/mm²]:": str(round(self._stresses[5], 2)),
                       "Preassure Rod [N/mm²]:": str(round(self._stresses[6], 2))}

        if self._n == 1:
            report_dict["Connection Type:"] = "single shear"

        if self._case == "Case 2":
            report_dict["Clamping Case:"] = "Case 2 (loose fit in rod and oversized fit in fork)"
        else:
            report_dict["Clamping Case:"] = "Case 3 (oversized fit in rod and loose fit in fork)"

        return report_dict

    def main(self):
        """
        Compute the bolt dimensions

        ...

        Parameters
        ----------
        None

        Returns
        -------
        None

        """

        # Calculate estimated, temporary diameter
        d_temp = self._calc_d_temp()

        # Defining Diameter in accordance to standard
        d_std = self._get_diameter(d_temp)

        # Check if the defined diameter is sufficient
        yieldstress_min = min(
            [self._mat_rod.yieldstress, self._mat_fork.yieldstress, self._bolt.material.yieldstress])
        d = self._check(d_std, yieldstress_min)

        self._bolt.diameter = d

        self._bolt.set_length(self._ws, self._n, self._t_r, self._t_f)

        self._bolt.set_dimensions(self._ws)

        self._connector.close_excel()

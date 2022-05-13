import getpass
import sys
import math
from datetime import datetime
from pathlib import Path
from tkinter import (Event, Frame, Grid, IntVar,
                     Label, Menu, Message, PhotoImage, Scrollbar, Text,
                     Toplevel, messagebox, ttk)
from typing import Dict, List

from AddMaterial import AddMaterial
from Bolt_DIN1445 import BoltDIN1445
from Bolt_ISO2340 import BoltISO2340
from Bolt_ISO2341 import BoltISO2341
from calculation import Calculation
from excel_connector import ExcelFileNotFoundError
from Info import Info
from Material import Material, MaterialFileError, MaterialFileNotFoundError
from Report import PathError, Report
from Settings import Settings
from solidworks_connector import (CADFileError, CADFileNotFoundError,
                                  CADFileWithSameNameAlreadyOpenError,
                                  SaveError, SldWorksConnectionError,
                                  SldWorksConnector, VariableError)


class PathFileNotFoundError(FileNotFoundError):
    """
    Custom FileNotFoundError. Raised when provided path to txt-file containing
    paths is erroneous.
    """

    def __init__(self, path: str) -> None:
        self._path = path
        self._message = f"The file with path:\n {self._path}\ncould not be opened."

        super().__init__(self._message)


class TextboxEmptyError(Exception):
    """
    Custom exception raised when a textbox is empty
    """

    def __init__(self, tb: ttk.Entry):
        self._name = tb.winfo_name().upper().replace("_", " ")
        self._message = f"A value for {self._name} must be defined."

        super().__init__(self._message)


class ComboboxNotSelectedError(Exception):
    """
    Custom Exception raised when no selcetion for combobox is made
    """

    def __init__(self, cb: ttk.Combobox) -> None:
        self._name = cb.winfo_name().upper().replace("_", " ")
        self._message = f"A selection for {self._name} must be made."

        super().__init__(self._message)


class MainWindow(Frame):
    """
    Class gor main window of bolt calculator
    """

    def __init__(self, master=None):
        Frame.__init__(self, master)

        # =====Loading Information============================================================================

        self._materials: List[Material] = list()
        self._mat_cb: List[str] = list()
        self._mat_nr_cb: List[str] = list()
        self._textboxes: List[ttk.Entry] = list()
        self._labels: List[Label] = list()
        self._comboboxes: List[ttk.Combobox] = list()
        self._paths: Dict[str, str] = dict()
        self._cwd: Path = self._get_cwd()
        self._direcotry_path: Path = self._cwd.joinpath(
            "docs", "directories.txt")

        self.master.title("Bolt Calculator")
        self.master.iconbitmap(self._cwd.joinpath("docs", "b_c_logo.ico"))
        self.master.resizable(False, False)
        self._get_directories()
        self._import_mat()
        # ====================================================================================================
        # =====Defining Frames================================================================================
        self._left_frame = Frame(master)
        self._left_frame.grid(row=0, column=0)

        self._right_frame = Frame(master)
        self._right_frame.grid(row=0, column=1)

        self._status_frame = Frame(
            master, highlightbackground="black", highlightthickness=0.5)
        self._status_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")

        self._image_frame = Frame(self._left_frame)
        self._image_frame.pack(fill="both")

        self._label_frame = Frame(self._left_frame)
        self._label_frame.pack(fill="both", expand=1)

        self._button_frame = Frame(self._left_frame)
        self._button_frame.pack(fill="both", expand=1)

        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self._statustext = f"{timestamp}>>>Bolt Calculator started"

        self._scrollbar = Scrollbar(self._status_frame, orient="vertical")
        self._scrollbar.pack(side="right", fill="y")

        self._status_bar = Text(
            self._status_frame, background="white", font=("Courier", 8), height=8, yscrollcommand=self._scrollbar.set)
        self._status_bar.pack(fill="both")
        self._status_bar.insert("end", self._statustext)

        self._scrollbar.config(command=self._status_bar.yview)

        # ====================================================================================================
        # =====Defining Menubar===============================================================================
        self._menubar = Menu(master)
        self.master.configure(menu=self._menubar)

        self._settings_menu = Menu(self._menubar, tearoff=0)
        self._settings_menu.add_command(
            label="Directories", command=self._settings)

        self._mat_IntVar = IntVar()
        self._mat_IntVar.set(1)
        self._material_menu = Menu(self._menubar, tearoff=0)
        self._material_menu.add_radiobutton(
            label="Materialname", value=1, variable=self._mat_IntVar, command=lambda: [self._setmat(), self._update_statusbar("Materials show with materialname")])
        self._material_menu.add_radiobutton(
            label="Materialnumber", value=2, variable=self._mat_IntVar, command=lambda: [self._setmat(), self._update_statusbar("Materials show with materialnumber")])

        self._force_IntVar = IntVar()
        self._force_IntVar.set(2)
        self._force_menu = Menu(self._menubar, tearoff=0)
        self._force_menu.add_radiobutton(
            label="Force components", value=1, variable=self._force_IntVar, command=lambda: [self._set_textboxes(), self._update_statusbar("Force property is set to Force Components")])
        self._force_menu.add_radiobutton(
            label="Resulting Force", value=2, variable=self._force_IntVar, command=lambda: [self._set_textboxes(), self._update_statusbar("Force property is set to Resulting Force")])

        self._data_menu = Menu(self._menubar, tearoff=0)
        self._data_menu.add_cascade(label="Forces", menu=self._force_menu)
        self._data_menu.add_cascade(label="Material", menu=self._material_menu)

        self._about_menu = Menu(self._menubar, tearoff=0)
        self._about_menu.add_command(label="About", command=self._about)

        self._menubar.add_cascade(label="Settings", menu=self._settings_menu)
        self._menubar.add_cascade(label="Data", menu=self._data_menu)
        self._menubar.add_cascade(label="About", menu=self._about_menu)
        # ====================================================================================================
        # =====Defining Layout Options========================================================================
        rows = len(self._labels)
        buttons = 3

        i = 0
        if i <= rows:
            Grid.rowconfigure(self._label_frame, index=i, weight=1)
            i += 1

        i = 0
        if i <= buttons:
            Grid.rowconfigure(self._button_frame, index=i, weight=1)
            i += 1

        Grid.columnconfigure(self._button_frame, index=0, weight=1)
        Grid.columnconfigure(self._button_frame, index=1, weight=1)
        Grid.columnconfigure(self._button_frame, index=2, weight=1)

        Grid.columnconfigure(self._label_frame, index=0, weight=1)
        Grid.columnconfigure(self._label_frame, index=1, weight=1)
        Grid.columnconfigure(self._label_frame, index=2, weight=1)
        Grid.columnconfigure(self._label_frame, index=3, weight=1)

        ttk.Style().configure("TLabel", font=("Calibri", 12))
        ttk.Style().configure("TEntry", font=("Calibri", 12))
        ttk.Style().configure("TRadiobutton", font=("Calibri", 12))
        # ====================================================================================================
        # =====Setting Labels=================================================================================
        self._lb_user = ttk.Label(
            self._image_frame, text="User ID:", font=("Calibri", 12, "bold"))
        self._lb_user.grid(row=0, column=0, padx=5, pady=2, sticky="nsew")

        self._lb_ID = ttk.Label(
            self._image_frame, text=getpass.getuser(), font=("Calibri", 12))
        self._lb_ID.grid(row=0, column=1, padx=5, pady=2, sticky="nsew")

        self._lb_PID = ttk.Label(
            self._image_frame, text="Part ID:", font=("Calibri", 12, "bold"))
        self._lb_PID.grid(row=0, column=2, padx=5, pady=2, sticky="nsew")

        self._lb_name = ttk.Label(
            self._image_frame, text="Bolt Name", font=("Calibri", 12, "bold"))
        self._lb_name.grid(row=0, column=4, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_name)

        self._lb_f1 = ttk.Label(self._label_frame, text="Force in 1 [N]")
        self._lb_f1.grid(row=0, column=0, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_f1)

        self._lb_f2 = ttk.Label(self._label_frame, text="Force in 2 [N]")
        self._lb_f2.grid(row=1, column=0, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_f2)

        self._lb_resulting_force = ttk.Label(
            self._label_frame, text="Resulting Force F_r [N]")
        self._lb_resulting_force.grid(
            row=2, column=0, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_f2)

        self._lb_load_type = ttk.Label(self._label_frame, text="Load Type")
        self._lb_load_type.grid(row=3, column=0, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_load_type)

        self._lb_safty = ttk.Label(self._label_frame, text="Safty Factor")
        self._lb_safty.grid(row=4, column=0, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_safty)

        self._lb_application_factor = ttk.Label(
            self._label_frame, text="Application Factor")
        self._lb_application_factor.grid(
            row=5, column=0, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_application_factor)

        self._lb_material = ttk.Label(self._label_frame, text="Material")
        self._lb_material.grid(row=6, column=0, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_material)

        self._lb_rod = ttk.Label(
            self._label_frame, text="Rod Thickness t_r [mm]")
        self._lb_rod.grid(row=0, column=2, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_rod)

        self._lb_rod_material = ttk.Label(
            self._label_frame, text="Rod Material")
        self._lb_rod_material.grid(
            row=1, column=2, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_rod_material)

        self._lb_fork = ttk.Label(
            self._label_frame, text="Fork Thickness t_f [mm]")
        self._lb_fork.grid(row=2, column=2, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_fork)

        self._lb_fork_material = ttk.Label(
            self._label_frame, text="Fork Material")
        self._lb_fork_material.grid(
            row=3, column=2, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_fork_material)

        self._lb_case = ttk.Label(self._label_frame, text="Clamping Case")
        self._lb_case.grid(row=4, column=2, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_case)

        self._lb_connection = ttk.Label(
            self._label_frame, text="Connection Type")
        self._lb_connection.grid(
            row=5, column=2, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_case)

        self._lb_standard = ttk.Label(self._label_frame, text="Standard")
        self._lb_standard.grid(row=6, column=2, padx=5, pady=2, sticky="nsew")
        self._labels.append(self._lb_standard)

        self._img = PhotoImage(file=self._cwd.joinpath("docs", "bolt.png"))
        self._lb_img = Label(self._right_frame,  image=self._img)
        self._lb_img.pack(padx=10)
        # ====================================================================================================
        # =====Setting Inputs ================================================================================
        self._tb_part_ID = ttk.Entry(
            self._image_frame, name="part ID", font=("Calibri", 12))
        self._tb_part_ID.grid(row=0, column=3, padx=5, pady=2, sticky="nsew")
        self._textboxes.append(self._tb_part_ID)
        # self._tb_part_ID.insert("end", "1234")

        self._tb_name = ttk.Entry(
            self._image_frame, name="name", font=("Calibri", 12))
        self._tb_name.grid(row=0, column=5, padx=5, pady=2, sticky="nsew")
        self._tb_name.bind("<FocusOut>", lambda e,
                           x=self._tb_name: self._on_changed(e, x))
        self._textboxes.append(self._tb_name)
        # self._tb_name.insert("end", "Test")

        self._tb_f1 = ttk.Entry(
            self._label_frame, name="f_1", width=24, state="disabled")
        self._tb_f1.grid(row=0, column=1, pady=2, padx=5, sticky="nsew")
        self._tb_f1.bind("<FocusOut>", lambda e,
                         x=self._tb_f1: self._on_changed(e, x))
        self._textboxes.append(self._tb_f1)

        self._tb_f2 = ttk.Entry(
            self._label_frame, name="f_2", width=24, state="disabled")
        self._tb_f2.grid(row=1, column=1, pady=2, padx=5, sticky="nsew")
        self._tb_f2.bind("<FocusOut>", lambda e,
                         x=self._tb_f2: self._on_changed(e, x))
        self._textboxes.append(self._tb_f2)

        self._tb_resulting_force = ttk.Entry(
            self._label_frame, name="f_r", width=24)
        self._tb_resulting_force.grid(
            row=2, column=1, pady=2, padx=5, sticky="nsew")
        self._tb_resulting_force.bind("<FocusOut>", lambda e,
                                      x=self._tb_resulting_force: self._on_changed(e, x))
        self._textboxes.append(self._tb_resulting_force)
        # self._tb_resulting_force.insert("end", "2000")

        self._cb_load_type = ttk.Combobox(self._label_frame, name="load type", state="readonly", values=[
            "static", "alternating", "pulsating"])
        self._cb_load_type.grid(row=3, column=1, pady=2, padx=5, sticky="nsew")
        self._comboboxes.append(self._cb_load_type)

        self._tb_safty = ttk.Entry(
            self._label_frame, name="safty factor", width=24)
        self._tb_safty.grid(row=4, column=1, pady=2, padx=5, sticky="nsew")
        self._tb_safty.bind("<FocusOut>", lambda e,
                            x=self._tb_safty: self._on_changed(e, x))
        self._textboxes.append(self._tb_safty)
        # self._tb_safty.insert("end", "1.5")

        self._tb_application_factor = ttk.Entry(
            self._label_frame, name="application factor", width=24)
        self._tb_application_factor.grid(
            row=5, column=1, pady=2, padx=5, sticky="nsew")
        self._tb_application_factor.bind("<FocusOut>", lambda e,
                                         x=self._tb_application_factor: self._on_changed(e, x))
        self._textboxes.append(self._tb_application_factor)
        # self._tb_application_factor.insert("end", "1")

        self._cb_material = ttk.Combobox(
            self._label_frame, name="material", state="readonly", values=self._mat_cb, width=24)
        self._cb_material.grid(row=6, column=1, pady=2, padx=5, sticky="nsew")
        self._comboboxes.append(self._cb_material)

        self._tb_rod = ttk.Entry(
            self._label_frame, name="rod thickness", width=24)
        self._tb_rod.grid(row=0, column=3, pady=2, padx=5, sticky="nsew")
        self._tb_rod.bind("<FocusOut>", lambda e,
                          x=self._tb_rod: self._on_changed(e, x))
        self._textboxes.append(self._tb_rod)
        # self._tb_rod.insert("end", "40")

        self._cb_rod_material = ttk.Combobox(
            self._label_frame, name="rod material", state="readonly", values=self._mat_cb, width=24)
        self._cb_rod_material.grid(
            row=1, column=3, pady=2, padx=5, sticky="nsew")
        self._comboboxes.append(self._cb_rod_material)

        self._tb_fork = ttk.Entry(
            self._label_frame, name="fork thickness", width=24)
        self._tb_fork.grid(row=2, column=3, pady=2, padx=5, sticky="nsew")
        self._tb_fork.bind("<FocusOut>", lambda e,
                           x=self._tb_fork: self._on_changed(e, x))
        self._textboxes.append(self._tb_fork)
        # self._tb_fork.insert("end", "10")

        self._cb_fork_material = ttk.Combobox(
            self._label_frame, name="fork material", state="readonly", values=self._mat_cb, width=24)
        self._cb_fork_material.grid(
            row=3, column=3, pady=2, padx=5, sticky="nsew")
        self._comboboxes.append(self._cb_fork_material)

        self._cb_case = ttk.Combobox(self._label_frame, name="clamping case", state="readonly", values=[
            "Case 1", "Case 2", "Case 3"], width=24)
        self._cb_case.grid(row=4, column=3, pady=2, padx=5, sticky="nsew")
        self._comboboxes.append(self._cb_case)

        self._cb_connection = ttk.Combobox(self._label_frame, name="connection type", state="readonly", values=[
            "single shear", "double shear"], width=24)
        self._cb_connection.grid(
            row=5, column=3, pady=2, padx=5, sticky="nsew")
        self._comboboxes.append(self._cb_connection)

        self._cb_std = ttk.Combobox(self._label_frame, name="standard", state="readonly", values=[
            "ISO 2340 A", "ISO 2340 B", "ISO 2341 A", "ISO 2341 B", "DIN 1445"], width=24)
        self._cb_std.grid(row=6, column=3, pady=2, padx=5, sticky="nsew")
        self._comboboxes.append(self._cb_std)
        # ====================================================================================================
        # =====Setting Buttons================================================================================
        self._btn_add = ttk.Button(self._button_frame, text="Add Material")
        self._btn_add.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self._btn_add.bind("<Return>", self._add_material)
        self._btn_add.bind("<Button-1>", self._add_material)

        self._btn_reset = ttk.Button(self._button_frame, text="Reset")
        self._btn_reset.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self._btn_reset.bind("<Return>", self._reset)
        self._btn_reset.bind("<Button-1>", self._reset)

        self._btn_calc = ttk.Button(self._button_frame, text="Calculate")
        self._btn_calc.grid(row=0, column=2, padx=10, pady=10, sticky="nsew")
        self._btn_calc.bind("<Return>", self._calc)
        self._btn_calc.bind("<Button-1>", self._calc)
        # ====================================================================================================

    def _get_cwd(self) -> Path:
        """
        Get the current working directory
        ...

        Parametes
        ---------
        None

        Returns
        -------
        cwd: Path
            current working directory
        """
        # determine if application is a script file or frozen exe
        if getattr(sys, 'frozen', False):
            return Path(sys.executable).parent.parent
        elif __file__:
            return Path(__file__).parent.parent

    def _settings(self) -> None:
        """
        Handle the opening and closing of the Settings window
        ...

        Parametes
        ---------
        None

        Returns
        -------
        None
        """
        # Toplevel() creates a new toplevel window
        settings = Settings(Toplevel(), self._cwd,
                            self._paths, self._direcotry_path, self._update_statusbar)
        self._paths = settings.to_main()
        self._get_directories()
        self._import_mat()
        self._setmat()

    def _about(self) -> None:
        """
        Handle the opening and closing of the About window
        ...

        Parametes
        ---------
        None

        Returns
        -------
        None
        """

        # Toplevel() creates a new toplevel window
        top_about = Toplevel()
        top_about.title("Bolt Calculator About")
        top_about.iconbitmap(self._cwd.joinpath("docs", "b_c_logo.ico"))
        top_about.resizable(False, False)
        text = Message(top_about, text=f"Bolt Caluclator Version: 1.1 \n\nInformation:\
            \nThis application uses third party software (Excel and SolidWorks) to calculate bolts and set their measurements in accordance to legal standards (DIN1445, ISO2340, ISO2341). It allows to create CAD modells in SolidWorks and calculation reports in PDF format. \
            \n\nDisclaimer: \
            \nThe creator of this application can not be held liable for possible failure of bolts calculated using this software or any legal consequences which might arise from such failures. \
            \n\nFor questions, suggestions or corrections please contact: \nyannick-keller@posteo.de")
        text.pack(padx=5, pady=5)

    def _get_directories(self) -> None:
        """
        Get directories from directories file
        ...

        Parameters
        ----------
        None

        Returns
        -------
        None
        """
        try:
            file = open(self._direcotry_path, "r", encoding="utf-8")
        except FileNotFoundError:
            raise PathFileNotFoundError(self._direcotry_path) from None

        for line in file:
            (key, val) = line.split(" = ")
            # .rstrip("\n") removes newline from string
            self._paths[key] = val.rstrip("\n")
        file.close()

    def _update_cb(self) -> None:
        """
        Update comboboxes with materialname or materialnumber representation
        of materials
        ...

        Parameters
        ----------
        None

        Returns
        -------
        None
        """
        for m in self._materials:
            self._mat_cb.append(m.to_string())
            self._mat_nr_cb.append(m.to_string_nr())

    def _update_statusbar(self, status: str) -> None:
        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self._status_bar.insert("end", f"\n{timestamp}>>>{status}")
        self._status_bar.see("end")
        # Causes Text to update immediatly
        self.update_idletasks()

    def _import_mat(self) -> None:
        """
        Import materialdata from .mat-file
        ...

        Parameters
        ----------
        None

        Returns
        -------
        None
        """

        try:
            self._materials = Material.mats_from_file(
                self._paths.get("material_directory"))
        except (MaterialFileNotFoundError, MaterialFileError) as error:
            messagebox.showerror(type(error).__name__, error.args[0])

        self._update_cb()

    def _reset(self, _: Event) -> None:
        """
        Reset all text- and comboboxes
        ...

        Parameters
        ----------
        _: Event
            the event triggering the function

        Returns
        -------
        None
        """
        for tb in self._textboxes:
            tb.configure(state="normal")
            tb.delete(0, "end")
        for cb in self._comboboxes:
            cb.set("")
        self._set_textboxes()

    def _check_textboxes(self) -> None:
        """
        Check if all textboxes are filled
        ...

        Parameters
        ----------
        None

        Returns
        -------
        None

        Raises
        ------
        TextboxEmptyError

        """

        for tb in self._textboxes:
            # if textbox is empty and activated raise an error
            # ttk.entry.stae() returns a tuple containing the states of the entry
            if len(tb.get()) == 0 and "disabled" not in tb.state():
                raise TextboxEmptyError(tb) from None

    def _check_comboboxes(self):
        """
        Check if a selection for each combobox is made
        ...

        Parameters
        ----------
        None 

        Returns
        -------
        None

        Raises
        ------
        ComboboxNotSelectedError
        """

        for cb in self._comboboxes:
            if len(cb.get()) == 0:
                raise ComboboxNotSelectedError(cb) from None

    def _calc(self, _: Event) -> None:
        """
        Start calculation
        ...

        Parameters
        ----------
        _: Event
            the event triggering the function

        Returns
        -------
        None
        """
        # Check if all text- and comboboxes are filled with valid values
        try:
            self._check_textboxes()
            self._check_comboboxes()
        except (TextboxEmptyError, ComboboxNotSelectedError) as error:
            self._update_statusbar(f"{type(error).__name__} occured")
            messagebox.showerror(type(error).__name__, error.args[0])
            return

        # set number of cutting surfaces according to connection type
        if self._cb_connection.get() == "single shear":
            n = 1
        else:
            n = 2

        self._update_statusbar("Calculation started")

        if self._cb_std.get() == "ISO 2340 A":
            bolt = BoltISO2340(self._tb_name.get(),
                               self._materials[self._cb_material.current()], "A")
        elif self._cb_std.get() == "ISO 2340 B":
            bolt = BoltISO2340(self._tb_name.get(),
                               self._materials[self._cb_material.current()], "B")
        elif self._cb_std.get() == "ISO 2341 A":
            bolt = BoltISO2341(self._tb_name.get(),
                               self._materials[self._cb_material.current()], "A")
        elif self._cb_std.get() == "ISO 2341 B":
            bolt = BoltISO2341(self._tb_name.get(),
                               self._materials[self._cb_material.current()], "B")
        else:
            bolt = BoltDIN1445(self._tb_name.get(),
                               self._materials[self._cb_material.current()])
        try:
            self._update_statusbar("Connecting to Excel...")
            calc: Calculation = Calculation(
                bolt,
                self._cb_load_type.get(),
                self._cb_case.get(),
                float(self._tb_resulting_force.get()),
                float(self._tb_rod.get()),
                float(self._tb_fork.get()),
                n,
                float(self._tb_application_factor.get()),
                float(self._tb_safty.get()),
                self._materials[self._cb_rod_material.current()],
                self._materials[self._cb_fork_material.current()],
                self._paths.get("measurements_input"))

            self._update_statusbar("Connection to Excel established")

        except ExcelFileNotFoundError as error:
            self._update_statusbar(f"{type(error).__name__} occured")
            messagebox.showerror(type(error).__name__, error.args[0])

        # Calculate bolt dimensions
        self._update_statusbar("Calculating...")
        calc.main()
        self._update_statusbar("Calculation terminated successfully")
        self._update_statusbar("Connection to Excel closed")

        top_info = Toplevel()
        info = Info(top_info, self._cwd, bolt)
        CAD, report = info.to_main()

        if CAD == 1:
            sw_connector = SldWorksConnector(
                self._paths.get("cad_input"), self._paths.get("cad_output"), bolt)
            try:
                sw_connector.main(self._update_statusbar)
            except (CADFileNotFoundError, CADFileError, CADFileWithSameNameAlreadyOpenError,
                    SldWorksConnectionError, VariableError, SaveError) as error:
                self._update_statusbar(f"{type(error).__name__} occured")
                messagebox.showerror(type(error).__name__, error.args[0])

        if report == 1:
            try:
                self._update_statusbar("Creating report...")
                report = Report(bolt, calc, self._paths.get(
                    "file_output"), self._update_statusbar)
                self._update_statusbar("Report created")
                self._update_statusbar("Report saved")
            except PathError as error:
                self._update_statusbar(f"{type(error).__name__} occured")
                messagebox.showerror(type(error).__name__, error.args[0])

    def _add_material(self, _: Event) -> None:
        """
        Open the Add Material window
        ...

        Parameters
        ----------
        _: Event
            event triggering the function

        Returns
        -------
        None
        """

        # creat new Toplevel window
        top_mat = Toplevel()
        # open add material window in the top level window
        add = AddMaterial(
            top_mat, self._cwd, self._paths["material_directory"], self._materials, self._update_statusbar)
        add.to_main()
        self._update_cb()
        self._setmat()

    def _on_changed(self, _: Event, textbox: ttk.Entry) -> None:
        """
        Check values entered into textboxes for validity
        ...

        Parameters
        ----------
        _: Event
            event triggering the function
        textbox: ttk.Entry
            textbox which is to be checked

        Returns
        -------
        None

        """
        # check if textbox is filled
        if len(textbox.get()) != 0:
            if textbox is self._tb_name:
                # check if value is numeric, if so show warning
                if textbox.get().isnumeric() is True:
                    self._update_statusbar("A Warning occured")
                    messagebox.showerror(
                        "Warning", "Name can not be a numeric value")
                    textbox.focus()
                    textbox.delete(0, "end")
                # check if name contains any unallowed characters, if so show warning
                else:
                    unallowed = ["!", "§", "$", "%", "&", "/", "@",
                                 "(", ")", "=", "?", "[", "]", ">", "<", ".", ",", ";", ":", "-"]
                    if any(char in unallowed for char in textbox.get()):
                        self._update_statusbar("A Warning occured")
                        messagebox.showerror(
                            "Warning", "Please use only letters, numbers, and underscores for the name")
                        textbox.delete(0, "end")

                    if " " in textbox.get():
                        name = textbox.get().replace(" ", "_")
                        self._tb_name.delete(0, "end")
                        self._tb_name.insert(0, name)
            else:
                try:
                    float(textbox.get())
                    # calculating the resulting force
                    if textbox is self._tb_f1 and len(self._tb_f2.get()) > 0:
                        F_r = str(round(math.sqrt(float(self._tb_f1.get())
                                                  ** 2 + float(self._tb_f2.get()) ** 2), 4))
                        self._tb_resulting_force.configure(state="normal")
                        self._tb_resulting_force.delete(0, "end")
                        self._tb_resulting_force.insert(0, F_r)
                        self._tb_resulting_force.configure(state="disabled")
                    elif textbox is self._tb_f2 and len(self._tb_f1.get()) > 0:
                        F_r = str(round(math.sqrt(float(self._tb_f1.get())
                                                  ** 2 + float(self._tb_f2.get()) ** 2), 4))
                        self._tb_resulting_force.configure(state="normal")
                        self._tb_resulting_force.delete(0, "end")
                        self._tb_resulting_force.insert(0, F_r)
                        self._tb_resulting_force.configure(state="disabled")
                except ValueError:
                    self._update_statusbar("ValueError occured")
                    messagebox.showerror(
                        "Warning", "Value must be numeric. Use . as decimal seperator")
                    textbox.delete(0, "end")
                    textbox.focus()

    def _setmat(self) -> None:
        """
        Update material comboboxes according to the selected option of how to
        show material information
        ...

        Parameters
        ----------
        None 

        Returns
        -------
        None
        """
        cur_mat = self._cb_material.current()
        cur_rod_mat = self._cb_rod_material.current()
        cur_fork_mat = self._cb_fork_material.current()
        self._cb_material.set("")
        if self._mat_IntVar.get() == 1:
            self._cb_material.config(values=self._mat_cb)
            self._cb_rod_material.config(values=self._mat_cb)
            self._cb_fork_material.config(values=self._mat_cb)
            if cur_mat >= 0:
                self._cb_material.set(self._mat_cb[cur_mat])
            if cur_rod_mat >= 0:
                self._cb_rod_material.set(self._mat_cb[cur_rod_mat])
            if cur_fork_mat >= 0:
                self._cb_fork_material.set(self._mat_cb[cur_fork_mat])

        elif self._mat_IntVar.get() == 2:
            self._cb_material.config(values=self._mat_nr_cb)
            self._cb_rod_material.config(values=self._mat_nr_cb)
            self._cb_fork_material.config(values=self._mat_nr_cb)
            if cur_mat >= 0:
                self._cb_material.set(self._mat_nr_cb[cur_mat])
            if cur_rod_mat >= 0:
                self._cb_rod_material.set(self._mat_nr_cb[cur_rod_mat])
            if cur_fork_mat >= 0:
                self._cb_fork_material.set(self._mat_nr_cb[cur_fork_mat])

    def _set_textboxes(self) -> None:
        """
        Activate or deactivate force comboboxes according to the selected 
        option for force input
        ...

        Parameters
        ----------
        None

        Returns
        -------
        None
        """
        if self._force_IntVar.get() == 1:
            self._tb_resulting_force.delete(0, "end")
            self._tb_resulting_force.configure(state="disabled")
            self._tb_f1.configure(state="normal")
            self._tb_f2.configure(state="normal")
        if self._force_IntVar.get() == 2:
            self._tb_resulting_force.configure(state="normal")
            self._tb_f1.delete(0, "end")
            self._tb_f2.delete(0, "end")
            self._tb_f1.configure(state="disabled")
            self._tb_f2.configure(state="disabled")


# master = Tk()
# app = MainWindow(master)
# app.mainloop()

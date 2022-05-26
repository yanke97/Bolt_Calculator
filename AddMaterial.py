from pathlib import Path
from tkinter import BOTH, END, Event, Frame, Toplevel, messagebox, ttk
from typing import Callable, List

from Material import (Material, MaterialFileNotFoundError, MaterialType,
                      MaterialTypeError)


class AddMaterial(Frame):
    """
    Class defining the AddMaterial window
    """

    def __init__(self, master=None, cwd: Path = None, path: str = None, materials: List[Material] = None, function: Callable = None):
        Frame.__init__(self, master)

        self.master: Toplevel = master
        self._cwd: Path = cwd
        self._path: str = path
        self._materials: List[Material] = materials
        self._function: Callable = function
        self._mattypes: List[MaterialType] = list()
        # List containing all textboxes in the GUI
        self._textboxes: List[ttk.Entry] = list()
        # List containing all labels in the GUI
        self._labels: List[ttk.Label] = list()

        self.master.title("Add Material")
        self.master.iconbitmap(self._cwd.joinpath("docs", "b_c_logo.ico"))
        self.master.resizable(False, False)
        self._import_mattypes()

        # Definign Frames for Layout of GUI
        self._label_frame = Frame(master)
        self._label_frame.pack(fill=BOTH, expand=1)

        self._button_frame = Frame(master)
        self._button_frame.pack(fill=BOTH, expand=1)

        # Defining style for label and entries
        ttk.Style().configure("TLabel", font=("Calibri", 12))
        ttk.Style().configure("TEntry", font=("Calibri", 12))

        self._lb_name = ttk.Label(self._label_frame, text="Name")
        self._lb_name.grid(row=0, column=0, pady=2, padx=10, sticky="nsew")
        self._labels.append(self._lb_name)

        self._lb_nr = ttk.Label(self._label_frame, text="Materialnr.")
        self._lb_nr.grid(row=1, column=0, pady=2, padx=10, sticky="nsew")
        self._labels.append(self._lb_nr)

        self._lb_density = ttk.Label(
            self._label_frame, text="Density [kg/m³]")
        self._lb_density.grid(row=2, column=0, pady=2, padx=10, sticky="nsew")
        self._labels.append(self._lb_density)

        self._lb_tensile_strength = ttk.Label(
            self._label_frame, text="Tensile Strength [N/mm²]")
        self._lb_tensile_strength.grid(
            row=3, column=0, pady=2, padx=10, sticky="nsew")
        self._labels.append(self._lb_tensile_strength)

        self._lb_yield_stress = ttk.Label(
            self._label_frame, text="Yield Stress [N/mm²]")
        self._lb_yield_stress.grid(
            row=4, column=0, pady=2, padx=10, sticky="nsew")
        self._labels.append(self._lb_yield_stress)

        self._lb_youngs_modulus = ttk.Label(
            self._label_frame, text="Youngs Modulus [N/mm²]")
        self._lb_youngs_modulus.grid(
            row=5, column=0, pady=2, padx=10, sticky="nsew")
        self._labels.append(self._lb_youngs_modulus)

        self._lb_type = ttk.Label(self._label_frame, text="Material Type")
        self._lb_type.grid(row=6, column=0, pady=2, padx=10, sticky="nsew")
        self._labels.append(self._lb_type)

        self._tb_name = ttk.Entry(self._label_frame, name="textbox_name")
        self._tb_name.grid(row=0, column=1, pady=2, padx=10, sticky="nsew")
        self._tb_name.bind("<FocusOut>", lambda e,
                           x=self._tb_name: self._on_changed(e, x))
        self._textboxes.append(self._tb_name)

        self._tb_nr = ttk.Entry(self._label_frame, name="textbox_nr")
        self._tb_nr.grid(row=1, column=1, pady=2, padx=10, sticky="nsew")
        self._tb_nr.bind("<FocusOut>", lambda e,
                         x=self._tb_nr: self._on_changed(e, x))
        self._textboxes.append(self._tb_nr)

        self._tb_density = ttk.Entry(
            self._label_frame, name="textbox_density")
        self._tb_density.grid(row=2, column=1, pady=2, padx=10, sticky="nsew")
        self._tb_density.bind("<FocusOut>", lambda e,
                              x=self._tb_density: self._on_changed(e, x))
        self._textboxes.append(self._tb_density)

        self._tb_tensile_strength = ttk.Entry(
            self._label_frame, name="textbox_tensileStrength")
        self._tb_tensile_strength.grid(
            row=3, column=1, pady=2, padx=10, sticky="nsew")
        self._tb_tensile_strength.bind(
            "<FocusOut>", lambda e, x=self._tb_tensile_strength: self._on_changed(e, x))
        self._textboxes.append(self._tb_tensile_strength)

        self._tb_yield_stress = ttk.Entry(
            self._label_frame, name="textbox_yieldStress")
        self._tb_yield_stress.grid(
            row=4, column=1, pady=2, padx=10, sticky="nsew")
        self._tb_yield_stress.bind(
            "<FocusOut>", lambda e, x=self._tb_yield_stress: self._on_changed(e, x))
        self._textboxes.append(self._tb_yield_stress)

        self._tb_youngs_modulus = ttk.Entry(
            self._label_frame, name="textbox_youngsModulus")
        self._tb_youngs_modulus.grid(
            row=5, column=1, pady=2, padx=10, sticky="nsew")
        self._tb_youngs_modulus.bind(
            "<FocusOut>", lambda e, x=self._tb_youngs_modulus: self._on_changed(e, x))

        self._cb_type = ttk.Combobox(
            self._label_frame, name="combobox_type", values=self._mattypes)
        self._cb_type.grid(row=6, column=1, pady=2, padx=10, sticky="nsew")

        self._btn_Add = ttk.Button(self._button_frame, text="Add")
        self._btn_Add.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self._btn_Add.bind("<Return>", self._add)
        self._btn_Add.bind("<Button-1>", self._add)

        self._btn_reset = ttk.Button(self._button_frame, text="Reset")
        self._btn_reset.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self._btn_reset.bind("<Return>", self._reset)
        self._btn_reset.bind("<Button-1>", self._reset)

        self._btn_cancel = ttk.Button(self._button_frame, text="Cancel")
        self._btn_cancel.grid(row=0, column=2, padx=10,
                              pady=10, sticky="nsew")
        self._btn_cancel.bind("<Return>", self._cancel)
        self._btn_cancel.bind("<Button-1>", self._cancel)

    def _reset(self, _: Event) -> None:
        """
        Clear all text- and comboboxes

        ...
        Parameters
        ----------
        event: Event
            event triggering the fucntion execution

        Retunrs
        -------
        None
        """

        for i in self._textboxes:
            i.delete(0, END)
        self._cb_type.set("")

    def _cancel(self, _: Event) -> None:
        """
        Close the AddMaterial Window. Defined Materials are not added

        ...
        Parameters
        ----------
        event: Event
            event triggering the fucntion execution

        Retunrs
        -------
        None
        """
        self._function("No material has been added")
        self.master.destroy()
        return 0

    def _add(self, _: Event) -> None:
        """
        Add new material to materiallist

        ...
        Parameters
        ----------
        event: Event
            event triggering the fucntion execution

        Retunrs
        -------
        None
        """

        filled = True
        for index, textbox in enumerate(self._textboxes):
            if len(textbox.get()) == 0:
                textbox.focus()
                filled = False
                label_text = self._labels[index]["text"]
                messagebox.showerror(
                    "Warning", f"Value for {label_text} must be defined")
                return

        if len(self._cb_type.get()) == 0:
            messagebox.showerror(
                "Warning", "Materialtype must be defined")
            self._cb_type.focus()
            filled = False

        if filled is True:
            # Check if definde material already exists
            if any((mat.name == self._tb_name.get() for mat in self._materials)) is True or any((mat.matnr == self._tb_nr.get() for mat in self._materials)) is True:
                messagebox.showinfo(
                    "Information", "The material you wanted to add is already specified.")

            else:
                try:
                    # Create Material Object
                    mat = Material(self._tb_name.get(), self._tb_nr.get(), float(self._tb_density.get()), float(self._tb_tensile_strength.get()),
                                   float(self._tb_yield_stress.get()), float(self._tb_youngs_modulus.get()), MaterialType[self._cb_type.get()])
                    # add material to .mat-file
                    Material.mat_to_file(self._path, mat)
                    self._materials.append(mat)
                    self._function("New material has been added succesfully.")
                except (MaterialTypeError, MaterialFileNotFoundError) as error:
                    self._function(f"{type(error).__name__} has occured.")
                    messagebox.showerror(
                        type(error).__name__, error.args[0])

                self.master.destroy()

    def _on_changed(self, _: Event, textbox: ttk.Entry) -> None:
        """
        Check if entrie to textboxes are valid

        Parameters
        ----------
        event: Event
            event triggering the fucntion execution

        textbox: ttk.Entry
            textbox to be checked

        Returns
        -------
        None
        """

        if textbox is self._tb_name:
            if len(textbox.get()) > 0:
                if textbox.get().isnumeric() is True:
                    messagebox.showerror(
                        "ValueError", "Name can not be a numeric value")
                    textbox.focus()
                    textbox.delete(0, END)
        else:
            if len(textbox.get()) > 0:
                try:
                    float(textbox.get())
                except ValueError as error:
                    messagebox.showerror(
                        type(error).__name__, "Value must be numeric. Use . as decimal seperator")
                    textbox.focus()
                    textbox.delete(0, END)

    def to_main(self) -> List[Material]:
        """
        Returning list of materials to main window after AddMaterial window
        is destroyed

        Parameters
        ----------
        None

        Returns
        -------
        materials: List(Material)
            list containing all materials
        """

        # Waits until AddMaterial window is destroyed
        self.wait_window()
        return self._materials

    def _import_mattypes(self) -> None:
        """
        Get Material Types for combobox

        ...
        Parameters
        ----------
        None

        Returns
        -------
        None
        """
        for mattype in MaterialType:
            self._mattypes.append(mattype.name)

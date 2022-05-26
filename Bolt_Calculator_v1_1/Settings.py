from tkinter import messagebox, ttk, filedialog, Toplevel, Frame, BOTH, END
from typing import Callable, List, Dict
from pathlib import Path


class PathFileNotFoundError(FileNotFoundError):
    """
    Custom FileNotFoundError. Raised when provided path to txt-file containing
    paths is erroneous.
    """

    def __init__(self, path: str) -> None:
        self._path = path
        self._message = f"The file with path:\n {self._path}\ncould not be opened."
        super().__init__(self._message)


class Settings(Frame):
    """
    Class defining the Settings window
    """

    def __init__(self, master=None, cwd: Path = None, paths: Dict[str, str] = None, directory_path: str = None, function: Callable = None):
        Frame.__init__(self, master)

        self.master: Toplevel = master
        self._labels: List = list()
        self._textboxes: Dict[str, ttk.Entry] = dict()
        self._paths: Dict[str, str] = paths
        self._directory_path: str = directory_path
        self._function: Callable = function

        self.master.title("Settings - Directories")
        self.master.iconbitmap(cwd.joinpath("docs", "b_c_logo.ico"))
        self.master.resizable(False, False)

        self._label_frame = Frame(master)
        self._label_frame.pack(fill=BOTH, expand=1)

        self._button_frame = Frame(master)
        self._button_frame.pack(fill=BOTH, expand=1)

        ttk.Style().configure("TLabel", font=("Calibri", 12))
        ttk.Style().configure("TEntry", font=("Calibri", 12))

        self._lb_mat_path = ttk.Label(self._label_frame, text="Material File")
        self._lb_mat_path.grid(row=0, column=0, pady=2, padx=5, sticky="nsew")
        self._labels.append(self._lb_mat_path)

        # self._lb_cad_in = ttk.Label(self._label_frame, text="CAD Input")
        # self._lb_cad_in.grid(row=1, column=0, pady=2, padx=5, sticky="nsew")
        # self._labels.append(self._lb_cad_in)

        self._lb_measurements_in = ttk.Label(
            self._label_frame, text="Measurements")
        self._lb_measurements_in.grid(
            row=2, column=0, pady=2,  padx=5, sticky="nsew")
        self._labels.append(self._lb_measurements_in)

        self._lb_file_out = ttk.Label(self._label_frame, text="Report Output")
        self._lb_file_out.grid(row=3, column=0, pady=2,  padx=5, sticky="nsew")
        self._labels.append(self._lb_file_out)

        self._lb_cad_out = ttk.Label(self._label_frame, text="CAD Output")
        self._lb_cad_out.grid(row=4, column=0, pady=2,  padx=5, sticky="nsew")
        self._labels.append(self._lb_cad_out)

        self._tb_mat_path = ttk.Entry(
            self._label_frame, name="material_directory", width=80)
        self._tb_mat_path.insert(0, self._paths.get("material_directory"))
        self._tb_mat_path.grid(row=0, column=1, pady=2, padx=5, sticky="nsew")
        self._textboxes["material_directory"] = self._tb_mat_path

        # self._tb_cad_in = ttk.Entry(self._label_frame, name="cad_input")
        # self._tb_cad_in.insert(0, self._paths.get("cad_input"))
        # self._tb_cad_in.grid(row=1, column=1, pady=2, padx=5, sticky="nsew")
        # self._textboxes["cad_input"] = self._tb_cad_in

        self._tb_measurements_in = ttk.Entry(
            self._label_frame, name="measurements_input")
        self._tb_measurements_in.insert(
            0, self._paths.get("measurements_input"))
        self._tb_measurements_in.grid(
            row=2, column=1, pady=2, padx=5, sticky="nsew")
        self._textboxes["measurements_input"] = self._tb_measurements_in

        self._tb_file_out = ttk.Entry(
            self._label_frame, name="file_output")
        self._tb_file_out.insert(0, self._paths.get("file_output"))
        self._tb_file_out.grid(row=3, column=1, pady=2, padx=5, sticky="nsew")
        self._textboxes["file_output"] = self._tb_file_out

        self._tb_cad_out = ttk.Entry(self._label_frame, name="cad_output")
        self._tb_cad_out.insert(0, self._paths.get("cad_output"))
        self._tb_cad_out.grid(row=4, column=1, pady=2, padx=5, sticky="nsew")
        self._textboxes["cad_output"] = self._tb_cad_out

        self._btn_browse_mat = ttk.Button(
            self._label_frame, text="Browse", command=lambda: self._browse_for_file(self._tb_mat_path))
        self._btn_browse_mat.grid(
            row=0, column=2, pady=2, padx=5, sticky="nsew")

        # self._btn_browse_cad_in = ttk.Button(
        #     self._label_frame, text="Browse", command=lambda: self._browse_for_directory(self._tb_cad_in))
        # self._btn_browse_cad_in.grid(
        #     row=1, column=2, pady=2, padx=5, sticky="nsew")

        self._btn_browse_measurements_in = ttk.Button(
            self._label_frame, text="Browse", command=lambda: self._browse_for_file(self._tb_measurements_in))
        self._btn_browse_measurements_in.grid(
            row=2, column=2, pady=2, padx=5, sticky="nsew")

        self._btn_browse_file = ttk.Button(
            self._label_frame, text="Browse", command=lambda: self._browse_for_directory(self._tb_file_out))
        self._btn_browse_file.grid(
            row=3, column=2, pady=2, padx=5, sticky="nsew")

        self._btn_browse_cad_out = ttk.Button(
            self._label_frame, text="Browse", command=lambda: self._browse_for_directory(self._tb_cad_out))
        self._btn_browse_cad_out.grid(
            row=4, column=2, pady=2, padx=5, sticky="nsew")

        self._btn_save = ttk.Button(
            self._button_frame, text="Save", command=self._save)
        self._btn_save.grid(row=0, column=0, pady=2, padx=5, sticky="nsew")

        self._btn_cancel = ttk.Button(
            self._button_frame, text="Cancel", command=self._cancel)
        self._btn_cancel.grid(row=0, column=1, pady=2, padx=5, sticky="nsew")

    def to_main(self) -> Dict[str, str]:
        """
        Returning dictionary of paths to main window after Settings window
        is destroyed

        Parameters
        ----------
        None

        Returns
        -------
        dict: Dict[str, str]
            dictionary containing all paths
        """

        # Waits until Settings window is destroyed
        self.wait_window()
        return self._paths

    def _browse_for_file(self, textbox: ttk.Entry) -> None:
        """
        Browse for file

        ...
        Parameters
        ----------
        textbox: ttk.Entry
            textbox to which the selected file-path should be written

        Retunrs
        -------
        None
        """

        if textbox.winfo_name() == "material_directory":
            path: str = filedialog.askopenfilename(
                initialdir="/", title="Select File",
                filetypes=([("mat-file", "*.mat")]))
        else:
            path: str = filedialog.askopenfilename(
                initialdir="/", title="Select File", filetypes=([("Excel-files", "*.xlsx")]))

        if len(path) > 0:
            textbox.delete(0, END)
            textbox.insert(0, path)

    def _browse_for_directory(self, textbox: ttk.Entry) -> None:
        """
        Browse for directory

        ...
        Parameters
        ----------
        textbox: ttk.Entry
            textbox to which the selected directory path should be written

        Returns
        -------
        None
        """

        path: str = filedialog.askdirectory(
            initialdir="/", title="Select Directory")

        if len(path) > 0:
            textbox.delete(0, END)
            textbox.insert(0, path)

    # Updates directory.txt (File in which directories are stored)
    def _update_directories(self) -> None:
        """
        Update .txt-file containing all paths

        ...
        Parameters
        ----------
        paths: Dict[str, str]
            dictionary containing all paths

        Returns
        -------
        None

        Raises
        ------
        PathFileNotFoundError
        """

        try:
            file = open(self._directory_path, "w",
                        newline="", encoding="utf-8")
        except FileNotFoundError:
            self._function("PathFileNotFoundError occured.")
            raise PathFileNotFoundError(self._directory_path) from None
        else:
            for key in self._paths:
                # \n starts a new line after each dict entry
                file.write(str(key) + " = " + self._paths[key] + "\n")
            file.close()

    def _save(self) -> None:
        """
        Save new directories

        ...
        Parameters
        ----------
        None

        Returns
        -------
        None
        """
        for key in self._paths:
            self._paths[key] = self._textboxes[key].get()

        try:
            self._update_directories()
            self._function("New directories saved successfully.")
        except PathFileNotFoundError as error:
            self._function("PathFileNotFoundError occured")
            messagebox.showerror(type(error).__name__, error.args[0])
        self.master.destroy()

    def _cancel(self) -> None:
        """
        Close the AddMaterial Window. Defined Materials are not added

        ...
        Parameters
        ----------
        event: Event
            event triggering the fucntion execution

        Returns
        -------
        None
        """
        self._function("Changed file paths have not been saved.")
        self.master.destroy()

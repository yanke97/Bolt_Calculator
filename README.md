# BoltCalculator
Tool to calculate bolts using the procedure stated in Roloff/ Matek 23rd edt. Dimensions of bolts will be set according to standards DIN1445, ISO2340 or ISO2341. Additionally, calculation reports or CAD models of the calculated bolts can be created. This tool uses third party software (Microsoft Excel/ SolidWorks). To perform calculations MS Excel is required, for the creation of CAD models SolidWorks is mandatory.

## Motivation
The aim of this tool is to support development engineers in their day-to-day work, by automating the frequently performed task of dimensioning bolts. This tool can help the engineers to focus on their main tasks and thereby decrease development lead time.

## Disclaimer
The creator of this application can not be held liable for possible failure of parts calculated using this software or any legal consequences which might arise from such failures. This project is licensed under the GNU AFFERO GENERAL PUBLIC LICENSE Version 3

## How to start
BoltCalculator does not require any installations, just download the "BoltCalculator.zip" here. Unzip it and execute the .exe-file found inside. Do not remove the .exe-file from the unzipped folder. You may create a shortcut to the .exe.

When first executing the .exe-file there might be an error message saying that some files could not be found. Nevertheless, BoltCalculator will still start, and you should be able to see the main window. Before performing any calculations, please set all directories to match where you stored BoltCalculator on your machine. If you close and restart BoltCalculator there should be no more error messages.

To perform calculations, you need to have Microsoft Excel installed on your machine. The CAD model creation currently only supports SolidWorks and therefore needs a working installation of SolidWorks

## User Guide
This section gives a brief overview of the textboxes and dropdown menus in the main window.

* **Force in 1/ Force in 2/ Resulting Force:** This textboxes allow you to specify the load applied to the bolt. You can either enter force components (Force in 1 and Force in 2) or a Resulting Force. If you choose to enter force components, the resulting force will be calculated automatically by the program. Default is Resulting Force, this can be changed by going to *Data → Forces*. All values must be entered in Newton.
* **Load Type:** In this dropdown menu specifies the load type (static, alternating or pulsating).
* **Safety Factor:** Factor added to increase the failure threshold of the bolt. Usually the factor takes a value between 1 and 2.
* **Application Factor:** Factor introduced to take into account load peaks caused by proper use of the bolt joint.
* **Material/ Rod Material/ Fork Material:** Dropdown menus to select materials for the bolt, fork and rod. BoltCalculator comes with a list of predefined materials. Additional materials can be added by clicking the *Add Material* button. Do not alter the .mat-file manually.
* **Rod Thickness/ Fork Thickness:** Values to describe the thickness of rod and bolt along the axis of the bolt. Values must be entered in Millimeter.
* **Clamping Case:** Case describing the fitting of the bolt in fork and rod. The cases are:
    * Case 1: Loose fit in fork and rod.
    * Case 2: Oversized fit in fork and loose fit in rod.
    * Case 3: Loose fit in fork and oversized fit in rod.
* **Connection Type:** Defining how many shear faces there are in the bolt joint. For a fork and a rod, it is *double shear*.
* **Standard:** Legal standard defining the bolt geometry.

## Contact
For corrections, suggestions and questions, please contact me via email: yannick-keller@posteo.de

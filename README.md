# Microglia Morphology Analysis


These Automated analysis of the morphology of microglia cells in H-DAB images using QuPath and ImageJ


////
//// Goals:   	- Given a DAB-stained-Image, detect all microglia soma body cells inside an annotated region and
////              analyse their morphology (number of branches, number of junctions...). This is done by 
////              first detecting the microglia cells in QuPath, and then sending a window around each cell
////              to ImageJ for further processing, skeletonization & skeleton analysis
////            - Send the resulting cell skeletons back to QuPath as an annotation overlay 
////            - Save morphology analysis results in an excel table & perform further statistical analysis
////
//// Required:  Brightfield Image with standard stain color vectors (DAB & Hematoxylin)
////            ImageJ Plugins: Skeletonization, Auto Local Thresholding (already included in the Fiji-version)
////
//// Results: This script creates 3 excel tables per run:
////          - <save_name>_full_skeletonization_results: The full, raw skeletonization results produced in ImageJ (including the number
////                                                       of branches, junctions, average branch length...). Each row represents the
////                                                       skeletonization results for a different, individual soma body cell.
////          - <save_name>_statistics_summary: Statistics summary file computed from the raw skeletonization results. Contains the mean value
////                                            for each column (e.g. the mean number of branches), as well as the annotation area measured
////                                            in pixels & microns, and the total number of detected soma body cells
////          - <save_name>_parameters: All functional parameters (cells_min_area, gaussian_sigma..) and their chosen values.
////                                    (This table can be used to later reproduce the results for any given analysis, and may also
////                                    be mentioned in the publication to give technical details on the analysis method used)
////


<div style="text-align: center;">
    <img src="results_figure.png" width="400"/>
    <p>TODO</p>
</div>

## Installation

1. Install QuPath (https://qupath.github.io/)
2. Install ImageJ / Fiji (https://imagej.net/software/fiji/downloads). Download the zipped folder, unzip folder, move Image-J folder to the desired directory, e.g. under C:/Program Files (x86)
3. Setup ImageJ Plugins in QuPath:
    1. Copy the path to the plugins directory of ImageJ (should look sth. like this: <Path-to-Fiji-installation>/Fiji/plugins)
    2. In the QuPath menu bar, select "Extensions" -> "ImageJ" -> "Set plugins directory", then choose the folder path of step 3.1
    3. Press "Select Folder", then close QuPath & re-open QuPath again (program needs to be closed to update the changes)
4. ImageJ should already contain all necessary plugins to run this script, but if some are missing, e.g. "Auto Local Threshold", do the following:
    1. Go to ImageJ Wiki (https://imagej.net/list-of-extensions) & search for the missing extension / plugin (e.g. Auto Local Threshold is found under: https://imagej.net/plugins/auto-local-threshold)
    2. Follow the installation instructions on the wiki page (There should be a link to download the plugins file. The downloaded file has to be copied to the  ImageJ-plugins directory, i.e. <Path-to-Fiji-installation>/Fiji/plugins)
5. The updated version of this script uses a Color Deconvolution plugin. Installation is as follows:
    1. Download the plugin from this github repository: https://github.com/landinig/IJ-Colour_Deconvolution2/blob/main/colour_deconvolution2.jar
    2. Place the downloaded file inside the ImageJ-plugins directory (<Path-to-Fiji-installation>/Fiji/plugins)
    3. You may have to close and re-open QuPath for the new plugins to be detected


## Usage Guide: Single Image Analysis
Use the script called `microglia_morphology_single_image.groovy` and follow these instructions if you only want to analyze a single image or to find good parameters for cell detection and thresholding before running the analysis on multiple images. Running the script in batch-mode for multiple images is explained
 [here](#usage-guide-analyze-multiple-images-batch-mode).

1. Open Project
    - **Case A) First time setup** (Setting up a new QuPath project)
        1. Create a new folder to store the QuPath project
        2. Open QuPath, then go to the "Project" tab in the left panel, then press "Create Project"
        3. Select the folder created in step 1.1
    - **Case B) Existing QuPath Project** (Script was used before on this machine)
        1. Open QuPath, then go to the "Project" tab in the left panel, then press "Open Project"
        2. Select your existing project folder, open this folder, then select the file called "project.qpproj" and press "open"

2. Add Images to Project
    1. Drag & Drop images into the left panel under "Project" (this will open another window, confirm by pressing "Import").
    2. A window with "Set image type" will pop up. Select "Brightfield H-DAB", then press "Apply"
    3. The images should then appear in this panel "Image List". An image can be opened by double-clicking on its name

3. Draw Annotation Region
    1. Zoom into the region of interest, then select one of the buttons with red forms (under the uppermost main menu bar). These are the annotation tools. The 6th from the left, the "brush tool", is especially useful (allows to freely draw)
    2. Draw annotation region (Having selected one of the drawing tool buttons, just left-click on the mouse to draw). Here are some tips:
        - **Change brush size:** The size of the brush adapts according to magnification, i.e. if you zoom in, the brush effectively paints small regions - while zoomed out it can quickly mark large areas
        - **Delete annotation:** Go to the "Annotations" tab, select the relevant annotation (should be highlighted in yellow), then press the "delete" button below the annotation list (or just press "delete" on the keyboard)
        - **Extend annotation:** Go to the "Annotations" tab, select the relevant annotation (highlighted in yellow). Click inside the annotation, then hold the left mouse button and draw to extend the region
        - **Erase parts of annotation:** Holding down the Alt-key while using the brush causes it to ‘subtract’ regions, basically acting as an eraser
        - For further info & tips, check the official [QuPath Annotation Documentation](https://qupath.readthedocs.io/en/0.3/docs/starting/annotating.html)

4. Open this script inside QuPath
    1. In the QuPath main menu, click "Automate" -> "Script Editor" to open the script editor.
    2. Inside the Script editor, press "File" -> "Open", then navigate to the script file location and select the script file (Alternatively, you can drag & drop the script file directly into the left panel ("Scripts") of the script editor)

5. Run script
    1. Select & open this script from the list of scripts in the left panel
    2. Choose a directory where all results should be saved:
        - Create a results folder, copy its path, and insert the path in line 117 (i.e., `excel_save_path = "<your-copied-path/results>"`)
        - (Note: If your copied path contains backslashes `\`, you need to replace them with forward slashes `/`)
    3. Choose a name under which the Excel table results for the current image should be saved:
        - Write your name in line 115, e.g., `excel_save_name = "Microglia_Img_1_dorsal_region"`
        - (IMPORTANT: Every time you run the script on a new annotated region in the same image, you should use a new name!)
    4. Optional: Adapt the functional parameters for the soma cell detection & ImageJ preprocessing, thresholding & skeletonization
        - e.g., if you also want to detect soma cells that are very small, decrease the cells' minimum area (e.g., `cells_min_area = 12.0`)
    5. Save changes: In the script editor menu, click "File" -> "Save"
    6. Select the annotation region in the image for which you want to run the script (if it is selected, it is highlighted in yellow)
    7. Press "run" (in the bottom right corner of the script editor)
    8. While the script is running, you can see what is currently computed by viewing the log in the lower part of the script editor. The last printed log is "INFO: DONE!", this signalizes that the script is done computing. You can now open & view the resulting Excel tables in your results folder (Note: You may have to close QuPath first in order to edit the Excel tables)


**Closing Remarks:**
- The script internally switches the DAB & Hematoxylin-channel to run the QuPath cell detection algorithm (since it is only implemented for the Hematoxylin channel, but we need to run it on the DAB-channel)
- If any errors occur or something is unclear, feel free to contact me under jume@di.ku.dk


## Usage Guide: Analyzing multiple images (batch mode)
Use the script `microglia_morphology_batchmode.groovy` if you want to analyze multiple images at once using QuPath's batch-mode. Note that it is advised to search for good parameters first (using the script for single image analysis decribed [above](#usage-guide-single-image-analysis)) before running the batchmode on multiple images.
The usage is almost identical to the single image analysis as decribed [above](#usage-guide-single-image-analysis), with some adjustments:

- You have to add all images to your QuPath project folder for which you want to run the analysis (see point 2 in the [single image analysis guide](#usage-guide-single-image-analysis))
- No name for the resulting excel file table needs to be given, you only need to provide the path to a folder where all results for all images should be saved (line 115: `excel_save_path = "<path-to-your-datafolder>/Results"`)
- The batchmode-version of the script expects pre-made annotations with a name for each image. This means that before running the script, you need to open each image and add it to your QuPath project, draw annotations for the regions you want to analyze, and then give these annotations a name. 
    - You can name annotations by right-clicking into the selected annotation, then pressing "Annotations" > "Set properties". This will open a window where you can type in a name; press "Apply" when you are done. 
- Remember to change the parameters in lines 119 - 145
"Microglia_Img["+ imageName +"]_region[" + annotation_dv.getName() + "]"
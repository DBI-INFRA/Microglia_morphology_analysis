//////////////////////////////////////////////////////////////////////////////////////////////////////////////
////
//// Name: 	    Microglia_morphology_analysis
//// Author: 	Julia Mertesdorf (DBI-INFRA IACF, jume@di.ku.dk)
//// Version: 	1.2
//// Date:      29/04/2024
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
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
////
//// INSTALLATION 
////
//// 1. Install QuPath (https://qupath.github.io/)
//// 2. Install ImageJ / Fiji (https://imagej.net/software/fiji/downloads)
////    (Download zipped folder, unzip folder, move Image-J folder to the desired directory, e.g. under C:/Program Files (x86))
//// 3. Setup ImageJ Plugins in QuPath:
////    3.1. Copy the path to the plugins directory of ImageJ (should look sth. like this: <Path-to-Fiji-installation>/Fiji/plugins)
////    3.2. In the QuPath menu bar, select "Extensions" -> "ImageJ" -> "Set plugins directory", then choose the folder path of step 3.1
////    3.3. Press "Select Folder", then close QuPath & re-open QuPath again (program needs to be closed to update the changes)
//// 4. Fiji-ImageJ should already contain all necessary plugins to run this script, but if some are missing,
////    e.g. "Auto Local Threshold", do the following:
////    4.1. Go to ImageJ Wiki (https://imagej.net/list-of-extensions) & search for the missing extension / plugin
////         (e.g. Auto Local Threshold is found under: https://imagej.net/plugins/auto-local-threshold)
////    4.2. Follow the installation instructions on the wiki page 
////         (There should be a link to download the plugins file. The downloaded file has to be copied to the 
////          ImageJ-plugins directory, i.e. <Path-to-Fiji-installation>/Fiji/plugins)
//// 5. The updated version of this script uses a Color Deconvolution plugin. Installation is as follows:
////    5.1. Download the plugin from this github repository: https://github.com/landinig/IJ-Colour_Deconvolution2/blob/main/colour_deconvolution2.jar
////    5.2. Place the downloaded file inside the ImageJ-plugins directory (<Path-to-Fiji-installation>/Fiji/plugins)
////    5.3. You may have to close and re-open QuPath for the new plugins to be detected
////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
////
//// USAGE GUIDE
////
//// 1. Open Project
////   1a. Case A) First time setup (Setting up a new QuPath project)
////       1a.1. Create a new folder to store the QuPath project
////       1a.2. Open QuPath, then go to Tab "Project" in the left panel, then press "Create Project"
////       1a.3. Select the folder created in step [1a.1]
////   1b. Case B) Existing QuPath Project (Script was used before on this machine)
////       1b.1. Open QuPath, then go to Tab "Project" in the left panel, then press "Open Project"
////       1b.2. Selct your existing project folder, open this folder, then select the file called "project.qpproj" & press "open"
//// 2.  Add Images to Project
////       2.1. Drag & Drop images into the left panel under "Project" (this will open another window, 
////            confirm by pressing "Import").
////       2.2. A window with "Set image type" will pop up. Select "Brightfield H-DAB", then press "Apply" 
////       2.3. The images should then appear in this panel "Image List". An image can be opened by double-clicking on its name
//// 3.  Draw Annotation Region
////       3.1. Zoom into the region of interest, then select one of the buttons with red forms (under the uppermost main menu bar).
////            These are the annotation tools. The 6th from the left, the "brush tool", is especially useful (allows to freely draw)
////       3.2. Draw annotation region (Having selected one of the drawing tool buttons, just left-click on mouse to draw).
////            Here are some tips:
////            - Change brush size: The size of the brush adapts according to magnification. In other words, if you zoom in, then
////                                 the brush effectively paints small regions - while zoomed out it can quickly mark in large areas
////            - Delete annotation: Go to "Annotations" tab, select the according annotation (should be highlighted in yellow),
////                                 then press on "delete" button below the annotation list (or just press "delete" on keyboard)
////            - Extend annotation: Go to "Annotations" tab, select the according annotation (highlighted in yellow).
////                                 Click inside the annotation, then hold the left mouse button and draw, to extend the region
////            - Erase parts of annotation: Holding down the Alt-key while using the brush causes it to ‘subtract’ regions, 
////                                         basically acting as an eraser
////            - For further info & tips, check the offical QuPath annotation documentation: 
////              https://qupath.readthedocs.io/en/0.3/docs/starting/annotating.html
//// 4. Open this script inside QuPath
////      4.1. In the QuPath main menu, click "Automate" -> "Script Editor". This opens the script editor.
////      4.2. Inside the Script editor, press "File" -> "Open", then navigate to the script file location & select the script file
////           (Alternatively, you can drag & drop the script file directly into the left panel ("Scripts") of the script editor)
//// 5. Run script
////      5.1. Select & open this script from the list of script in the left panel
////      5.2. Choose a directory where all results should be saved:
////           Create a results folder, copy its path & insert the path in line 113 (i.e. excel_save_path = "<your-copied-path/results>")
////           (Note: If your copied path contains backslashes \, you need to replace them by frontslashes /)
////      5.3. Choose a name under which the excel table results for the current image should be saved:
////           Write your name in line 111, e.g.: excel_save_name = "Microglia_Img_1_dorsal"
////           (IMPORTANT: Every time you run the script on a new annotated region in the same image, you should use a new name!)
////      5.4. Optional: Adapt the functional parameters for the soma cell detection & ImageJ preprocessing, thresholding & skeletonization
////           - e.g. if you also want to detect soma cells that are smaller, decrease the cells minimum area (e.g. cells_min_area = 18.0)
////      5.5. Save changes: In the script editor menu, click "File" -> "Save"
////      5.6. Select the annotation region in the image for which you want to run the script (if it is selected, it is highlighted in yellow)
////      5.7. Press "run" (in the bottom right corner of the script editor)
////      5.8. While the script is running, you can see what is currently computed by viewing the log in the lower part of the script editor. 
////           The last printed log is "INFO: DONE!", i.e. when you see this, the script is done computing. You can now open & view the resulting 
////           excel tables in your results folder (Note: You may have to close QuPath first in order to edit the Excel-tables)
////
//// CLOSING REMARKS: - The script internally switches the DAB & Hematoxylin-channel in order to be able to run the QuPath 
////                    cell detection algorithm (only implemented for Hematoxylin channel, but we need to run it on DAB-channel)
////                  - If any errors occur or something is unclear, feel free to contact me (jume@di.ku.dk)
////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////


// #################################### SAVING PARAMETERS #####################################################

// Define excel table result name & file location
excel_save_name = "Microglia_Img[18.59.20]_region[dorsal]" // Name that's used for saving the excel table 
                                                           // (Note: Use a different name for each image & annotation! Otherwise, the results from previous image calculations will be overwritten!)
excel_save_path = "<path-to-your-datafolder>/Results"  // Path to directory were the resulting csv files should be saved


// ################################### FUNCTIONAL PARAMETERS ##################################################

// 1) Functional parameters for soma body cell detection in QuPath
cells_min_area = 19.0         // Minimum area (in square microns) that an object must occupy to be considered a soma cell detection
cells_min_intensity = 0.1     // Intensity threshold for segmenting soma cells in the image (pixel intensity of the DAB-channel)
cells_min_circularity = 0.35  // Required minimum circularity to be considered a soma body cell
cells_min_caliper = 3.5       // The minimum diameter of soma body cell in pixels (all below this threshold are filtered out)
cells_min_dab_OD_mean = 0.15  // The minimum required mean optical density (OD), i.e. average intensity of DAB staining within a detected cell

window_width = 300            // Width of the rectangle annotation that is drawn around soma cell & sent to ImageJ [Recommended: 300]
window_height = 300           // Height of the rectangle annotation that is drawn around soma cell & sent to ImageJ [Recommended: 300]
                              // (Note: Width & Height should be set to the maximum possible size a soma cell and all of its branches can occupy)
cell_overlap_threshold = 0.75 // Defines the overlap threshold at which two detected cells are considered to be the same cell.
                              //  E.g. if the threshold is set to 0.8, it means that if the rectangles drawn around two cells overlap
                              //  to at least 80%, they are considered the same cell and are merged. If the overlap is less than 80%,
                              //  they are considered to be individual, distincet soma cells that are just very close to each other.

// 2) Functional parameters for branch morphology computation in ImageJ
threshold_variant = "global"        // Select the threshold type that should be used during the ImageJ skeletonization macro script 
                                    // Options: "global", "local" ("global" performs global thresholding, while "local" performs local, adaptive thresholding depending on a given radius)
                                    // - Note: Depending on whether "global" or local" is selected, this script will choose the global or local threshold type accordingly (defined in line 135 & 136)
global_threshold_type = "Triangle"  // Method used for auto global thresholding (Recommended: "Triangle", "MinError(I)", maybe also try: "Li", Otsu", "Moments")
local_threshold_type = "Niblack"    // Method used for auto local, adaptive thresholding (Recommended: "Niblack", "Otsu", try also "Bernsen" with radius ~10)
local_threshold_radius = 15         // The radius used during auto local thresholding (a smaller value will catch finer structures (branches), but may also catch more noise)

gaussian_sigma = 2.0                // Gaussian blur standard deviation (defines the amount of smoothing applied to the image)
bp_filter_large = 40                // Size (in pixels) of large bandpass filter (higher values retain low-frequency details like background & large features, while lower values remove them, enhancing smaller details)
bp_filter_small = 3                 // Size (in pixels) of small bandpass filter (smaller values retain high-level frequencies like smaller, detailled features in the image)
downsample_factor = 1               // Downsampling factor applied when sending a region from QuPath to ImageJ (1 = original resolution (recommended), 2 = resolution is halved)


// ######################################## PACKAGE IMPORTS ###################################################

import qupath.imagej.gui.ImageJMacroRunner
import qupath.lib.color.ColorDeconvolutionStains
import java.io.*
import java.io.File
import java.io.FileWriter
import java.io.BufferedWriter
import java.time.LocalDate
import java.time.format.DateTimeFormatter


// ############################################################################################################
// ########################################## MAIN SCRIPT #####################################################

println "[Step 1 / 5] ----------------------------- PREPROCESSING ------------------------------------------------"

// Create new folder to save all resulting csv-files to
(results_path, files_found) = create_folder(excel_save_path, excel_save_name)
if (files_found) {
    println "ERROR: Your previous results would have been overwritten - Please specify " +
             "a new name for the excel files, or delete the old files."
   return
}
clear_results_directory(results_path + "/intermediate_results", false)

// Set up logging to a separate log-file saved in the results-folder
logFilePath = results_path + "/logfile.log"
fileOutputStream = new FileOutputStream(logFilePath)   // Create a FileOutputStream to write to the log file
logPrintStream = new PrintStream(fileOutputStream)     // Create a PrintStream to write to the FileOutputStream

def customPrintStream = new PrintStream(System.out) {  // Wrap original System.out in a custom PrintStream that duplicates output to console & file
    @Override
    void write(byte[] buf, int off, int len) {
        super.write(buf, off, len)                     // Write to console
        fileOutputStream.write(buf, off, len)          // Write to log file
    }
}

// Redirect the console output to the custom PrintStream
System.setOut(customPrintStream)

// Save the used parmaters to an excel table in the same results folder (for reproducibility)
params_dict = create_params_dict(cells_min_area, cells_min_intensity, cells_min_circularity, cells_min_caliper, cells_min_dab_OD_mean, window_width, \
                                 window_height, cell_overlap_threshold, threshold_variant, global_threshold_type, local_threshold_type, \
                                 local_threshold_radius, gaussian_sigma, bp_filter_large, bp_filter_small, downsample_factor)
create_parameters_excel_table(params_dict, results_path, excel_save_name)

// Select current image to run script on
imageData = getCurrentImageData()

// Clear previous annotations & detections
clear_rectangular_annotations()
delete_all_detection_objects()
delete_all_skeleton_annotations()

// Swap the stain vectors DAB & Hematoxylin (to be able to run cell detection on DAB-channel)
set_stain_vectors_swapped()

// Check if an annotation region was drawn
annotation_exists = check_if_annotation_exists()
if (!(annotation_exists)) {
    println "ERROR: No annotation detected. Draw an annotation region first.\n       PROGRAM ABORTED"
    return
}
setup_annotation_regions()

// Get current selected annotation region & raise error if none selected
selected_annotation = getCurrentHierarchy().getSelectionModel().getSelectedObject()
if (selected_annotation == null) {
    println "ERROR: You have to select an annotation region first before running this script.\n       PROGRAM ABORTED"
    return
}
// Measure size of currently selected annotation region
(area_pixels, area_microns) = measure_annotation_region(selected_annotation)


println "\n      [Step 2 / 5] ----------------------------- SOMA CELL DETECTION & FILTERING ------------------------------"

// Run cell detection algorithm with parameters specified in lines 6ff
detect_soma_body_cells(cells_min_area, cells_min_intensity, cells_min_circularity, cells_min_caliper, cells_min_dab_OD_mean)

// Check if any soma body cells were found - if not, exit program
if (getDetectionObjects().isEmpty()) {
   println "Error: No soma body cells found. Try to annotate another or bigger region.\n      PROGRAM ABORTED"
   return
}

// Iterate over all detected cells & compute rectangle area around each cell
all_cell_windows = []
for (cell in getDetectionObjects()) {
    cell_roi = cell.getROI()  // Get the ROI (Region of Interest) associated with the cell
    window = get_window_around_cell(cell_roi, window_width, window_height)  // Draw window around cell & save this in the list of all cell windows

    // Check if the currently processed cell is a new cell (i.e. doesn´t overlap with cells already processed)
    new_soma_cell = check_if_new_cell(all_cell_windows, window, cell_overlap_threshold)
    if (new_soma_cell) {
        all_cell_windows.add(window)   // Only add cell window if it is a new, different cell
    }
}
number_of_soma_cell_detections = all_cell_windows.size()


println "\n      [Step 3 / 5] ----------------------------- SKELETONIZATION & QUANTIFICATION -----------------------------"

macro_params = instantiate_imageJ_macro(downsample_factor)  // Initialize ImageJ Macro-Runner

// Define an annotation class for the newly generated skeleton overlay annotations
def skeletonAnnotationClass = PathClass.getInstance("microglia_skeleton")
skeletonAnnotationClass.setColor(0, 255, 0)

// Pre-check that there are no errors in macro construction
(skeletonize_macro, error) = skeletonization_analysis_macro(results_path + "/intermediate_results", 1, gaussian_sigma, bp_filter_large, bp_filter_small,
                                                            150, threshold_variant, global_threshold_type, local_threshold_type, local_threshold_radius)
if (error) {
    println "Error: Your chosen method for ImageJ preprocessing called '${threshold_variant}' does not exist. Please select from ['local', 'global']\n      PROGRAM ABORTED"
    return
}

// Iterate over windows, send each to Image J for further processing & skeletonization
all_annotations_before = getAnnotationObjects()
increasing_anno_list = getAnnotationObjects()
floodfill_pos = (window_width / downsample_factor) / 2
cell_counter = 1
all_cell_windows.eachWithIndex { window, index ->
    println "Processing cell detection [${index + 1} / ${number_of_soma_cell_detections}]"
    csv_nr = String.format("%04d", index + 1)
    (skeletonize_macro, error) = skeletonization_analysis_macro(results_path + "/intermediate_results", csv_nr, gaussian_sigma, \
                                                                bp_filter_large, bp_filter_small, floodfill_pos, threshold_variant, \
                                                                global_threshold_type, local_threshold_type, local_threshold_radius)
    ImageJMacroRunner.runMacro(macro_params, imageData, null, window, skeletonize_macro)
    Thread.sleep(30)
    all_annos = getAnnotationObjects()
    new_anno = get_newest_annotation(increasing_anno_list, all_annos)
    if (new_anno != null) {
        new_anno.setName("${cell_counter}")
        cell_counter++
        increasing_anno_list = all_annos
    }
}

// Get the list of all newly added microglia skeleton annotations & set them to the according class (to get overlay)
all_annotations_after = getAnnotationObjects()

microglia_skeleton_overlay = []
for (anno in all_annotations_after) {
   if (!(anno in all_annotations_before)) {
       microglia_skeleton_overlay.add(anno)
       anno.setPathClass(skeletonAnnotationClass)
   }
}

println "\n      [Step 4 / 5] ----------------------------- EXCEL FILE MERGING -------------------------------------------"

(merged_csv_path, invalid_soma_cells) = merge_csv_files(results_path, excel_save_name)
if (invalid_soma_cells != 0) {
    println "Note: For ${invalid_soma_cells} detected soma cell(s), no skeleton could be computed." +
            "\n      These soma cells will not be considered for the statistics summary"
}
corrected_cell_counter = cell_counter - invalid_soma_cells - 1
compute_summary_csv(merged_csv_path, results_path, excel_save_name, area_pixels, area_microns, corrected_cell_counter)

println "\n      [Step 5 / 5] ----------------------------- POSTPROCESSING -----------------------------------------------"

// Remove rectangular window annotations (to declutter the overlay)
clear_rectangular_annotations()

clear_results_directory(results_path + "/intermediate_results", true)

println "DONE!"


// ############################################################################################################
// ##################################### HELPER FUNCTIONS FOR MAIN SCRIPT #####################################

// ===================================== FOLDER & EXCEL FILE PROCESSING =======================================

// Function to create a new subfolder at the specified path, if it doesn´t exist already
def create_folder(folder_path, excel_name) {
    current_date = LocalDate.now()
    date_formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd")
    formatted_date = current_date.format(date_formatter)
    
    folder_name = folder_path + "/" + formatted_date + "_" + excel_name
    folder = new File(folder_name)
    files = folder.listFiles()
    files_found = false
    if (files && files.length > 0) {
        println "There are already files saved under: ${folder_name}"
        files_found = true
    }
    folder_name_interm = folder_name + "/intermediate_results"
    folder_interm = new File(folder_name_interm)
    if (!folder_interm.exists()) {
        folder_interm.mkdirs()
    }
    return [folder_name, files_found]
}

// Function to delete all previous csv result files stored in directory "dir_path"
def clear_results_directory(dir_path, delete_folder) {
    folder = new File(dir_path)
    if (folder.exists() && folder.isDirectory()) {
        folder.eachFile { file ->
            if (file.isFile()) {
                file.delete()
            }
        }
        if (delete_folder) {
           println "Deleted folder '${dir_path}' successfully." 
           folder.deleteDir()
        }
        println "All files previously stored in '${dir_path}' deleted successfully."
    } else {
        println "ERROR: Specified folder does not exist or is not a directory: '${dir_path}'"
    }
}

// Function to create a dictionary saving the paramter values for all changeable parameters.
def create_params_dict(cells_min_area, cells_min_intensity, cells_min_circularity, cells_min_caliper, cells_min_dab_OD_mean, window_width, \
                       window_height, cell_overlap_threshold, threshold_variant, global_threshold_type, local_threshold_type, \
                       local_threshold_radius, gaussian_sigma, bp_filter_large, bp_filter_small, downsample_factor) {
    params_dict = [:]
    params_dict['cells_min_area'] = cells_min_area
    params_dict['cells_min_intensity'] = cells_min_intensity
    params_dict['cells_min_circularity'] = cells_min_circularity
    params_dict['cells_min_caliper'] = cells_min_caliper
    params_dict['cells_min_dab_OD_mean'] = cells_min_dab_OD_mean
    params_dict['window_width'] = window_width
    params_dict['window_height'] = window_height
    params_dict['cell_overlap_threshold'] = cell_overlap_threshold
    params_dict['threshold_variant'] = threshold_variant
    params_dict['global_threshold_type'] = global_threshold_type
    params_dict['local_threshold_type'] = local_threshold_type
    params_dict['local_threshold_radius'] = local_threshold_radius
    params_dict['gaussian_sigma'] = gaussian_sigma
    params_dict['bandpass_filter_large'] = bp_filter_large
    params_dict['bandpass_filter_small'] = bp_filter_small
    params_dict['downsample_factor'] = downsample_factor
    //println "${params_dict}"
    return params_dict
}

// Function to save the parameters dictionary as an excel table (to be able to reproduce the results later)
def create_parameters_excel_table(params_dict, folder_path, save_name) {
    csvFilePath = folder_path + "/" + save_name + "__parameters.csv"
    writer = new FileWriter(csvFilePath)
    keysString = params_dict.keySet().join(',')
    valuesString = params_dict.values().join(',')
    writer.write(keysString + "\n")
    writer.write(valuesString + '\n')
    writer.close()
    println "Parameters excel file saved successfully at: $csvFilePath"
}

// Function to merge all individual csv result files (containing the branching information) into one big csv file
def merge_csv_files(folder_path, file_name) {
    // Specify paths to new generated csv summary file & intermediate csv result files
    merged_csv_file_path = folder_path + "/" + file_name + "__full_skeletonization_results.csv"
    folder = new File(folder_path + "/intermediate_results")
    
    // Create a FileWriter and BufferedWriter to write to the new csv summary file
    writer = new FileWriter(merged_csv_file_path)
    
    // Check if folder path is valid, then iterate over all found intermediate csv result files
    invalid_soma_cells = 0
    if (folder.exists() && folder.isDirectory()) {
        try {
            csv_files = folder.listFiles()
            read_first_file = false
            csv_files.eachWithIndex { file, index ->
                csv_file = new File(file.path)
                
                // Check if the file is not empty
                lines = csv_file.readLines()
                lines = lines.findAll { line -> line.trim().length() > 0 }
                if (!(lines.isEmpty())) {
                    // For the first csv file, read & write the first line containing the branch statistics names
                    if (!read_first_file) {
                        writer.write("ID,")
                        line = csv_file.readLines().get(0)
                        writer.write(line.substring(line.indexOf(',') + 1))
                        writer.append('\n')
                        read_first_file = true
                    }
        
                    // Read results line from the current CSV file & write the line to the new merged csv file
                    line = csv_file.readLines().get(1)
                    number_of_entries = line.split(",").size()
                    if (number_of_entries == 10) {
                        writer.write("${index + 1},") // Add unique ID to each row
                        writer.write(line.substring(line.indexOf(',') + 1))
                        writer.append('\n')
                    }
                    else {
                        println "ERROR: ImageJ saved the wrong table - do not save branching information of cell ${index + 1}"
                        invalid_soma_cells += 1
                    }
                    
               } else {
                    invalid_soma_cells += 1
                }
            }
        } catch (IOException e) {
            println "ERROR: Error writing to CSV file: ${e.message}"
        } finally {
            try {
                writer.close()
            } catch (IOException e) {
                println "ERROR: Error closing FileWriter: ${e.message}"
            }
        }
    } else {
        println "ERROR: Merge csv files failed: Folder does not exist or is not a directory."
    }
    merged_csv_file = new File(merged_csv_file_path)
    merged_csv_file.setWritable(true)
    println "New excel file with full skeletonization results saved at: ${merged_csv_file_path}"
    return [merged_csv_file_path, invalid_soma_cells]
}

// Function to create the final csv file containing the summarized statistics about the microglia morphology
def compute_summary_csv(csv_path, output_dir, file_name, area_pixels, area_microns, detected_cells) {
    // Read the input CSV file
    merged_csv = new File(csv_path)
    if (!merged_csv.exists()) {
        println "ERROR: Merged csv file does not exist at: ${csv_path}"
        return
    } 

    // Create a dictionary for the csv data & Read the CSV file
    csv_data = [:]  // Dictionary
    csv_reader = new BufferedReader(new FileReader(merged_csv))
    header = csv_reader.readLine()?.split(",")   // Read the header row
    modified_header = header.collect { columnName -> "Mean ${columnName}" }  // Create adapted header

    // Create keys in the dictionary for each header column
    modified_header.each { columnName ->
        csv_data[columnName] = []
    }

    // Read data from each row & store it in the according dictionary list
    String row
    while ((row = csv_reader.readLine()) != null) {
        row_content = row.split(",")
        modified_header.eachWithIndex { modifiedColumnName, index ->
            csv_data[modifiedColumnName] << row_content[index].toDouble()
        }
    }
    csv_reader.close()
    //println "${csv_data}"

    // Create a new dict that stores the mean data of each column
    mean_data = [:]
    modified_header.each { columnName ->
        columnData = csv_data[columnName]
        sum = columnData.sum()
        mean_data[columnName] = (sum / columnData.size()).round(3)
    }
    // Add annotation area measurements
    modified_header += "Annotation Area (pixels)"
    modified_header += "Annotation Area (microns)"
    mean_data["Annotation Area (pixels)"] = area_pixels.round(3)
    mean_data["Annotation Area (microns)"] = area_microns.round(3)
    
    // Delete unused first column & add number of detected soma body cells
    mean_data.remove("Mean  ")
    modified_header = modified_header.drop(1)
    modified_header += "Total number of soma cells"
    mean_data["Total number of soma cells"] = detected_cells

    // Save results to a new CSV file
    statistics_summary_path = output_dir + "/" + file_name + "__statistics_summary.csv"
    writer = new FileWriter(statistics_summary_path)
    writer.write(modified_header.join(',') + '\n')  // Write header row
    writer.write(modified_header.collect { mean_data[it].toString() }.join(',') + '\n')  // Write mean data row
    writer.close()

    println "New excel file with statistics summary saved at: ${statistics_summary_path}"
}


// ===================================== ANNOTATION PROCESSING =============================================

// Function to clear all rectangular annotations
def clear_rectangular_annotations() {
    println "Deleting all rectangular annotations"
    // Get all the annotations in the image
    annotations = getAnnotationObjects() //  getAnnotationObjects().getObjects().collect() // getAnnotationObjects()
    
    // Iterate through each annotation
    for (annotation in annotations) {
        shape = annotation.getROI()
        if (shape && (shape.getRoiName() == "Rectangle")) {
            // Remove the rectangular annotation
            //getAnnotationObjects().removeObject(annotation)
            removeObject(annotation, true)
        }
    } 
}

// Function to clear all detection objects
def delete_all_detection_objects() {
    // Get all the detection objects in the image
    detections = getDetectionObjects()

    // Iterate through each detection object and remove it
    detections.each { detectionObject -> removeObject(detectionObject, true) }
}

// Function to delete all microglia skeleton annotations 
def delete_all_skeleton_annotations() {
    println "Deleting all microglia skeleton annotations"
    annotations = getAnnotationObjects()
    
    // Iterate through each annotation
    for (annotation in annotations) {
        anno_class = annotation.getPathClass() 
        if (anno_class != null && anno_class.getName() == "microglia_skeleton"){
            removeObject(annotation, true)
        }
    }
}

// Function to check if any annotations exist
def check_if_annotation_exists() {
    all_annotations = getAnnotationObjects()
    if (!(all_annotations)) {
        return false
    }
    return true
}

// Function to create & set the class of all original annotation regions
def setup_annotation_regions() {
    def originalAnnotationRegion = PathClass.getInstance("original_annotation_region")  // Define an annotation class
    all_annotations = getAnnotationObjects()
    for (anno in all_annotations) {
        anno.setPathClass(originalAnnotationRegion)
    }
    println "Successfully set type of all original annotation regions"
}

// Function to find the most recently added annotation by comparing two annotation lists
def get_newest_annotation(before_list, after_list) {
    //println "BEFORE: ${before_list.size()}, AFTER: ${after_list.size()}"
    newest_anno = null
    new_annos = 0
    for (anno in after_list) {
        id = anno.getID()
        if (!(before_list.find { it.getID() == id })) {
           newest_anno = anno
           new_annos += 1
       }
    }
    if (new_annos > 1 || newest_anno == null) {
        println "ERROR: NONE OR MORE THAN ONE NEW ANNOTATION OBJECT FOUND!"
    }
    return newest_anno
}

// Function to create a class for the original annotation region & measure its area size
def measure_annotation_region(selected_annotation) {
    pixel_size_microns = imageData.getServer().getPixelCalibration().getAveragedPixelSizeMicrons()
    area_pixels = selected_annotation.getROI().getArea()
    area_microns_squared = area_pixels * pixel_size_microns * pixel_size_microns
    println "Size of original annotation region: ${area_pixels} (pixels), ${area_microns_squared} (um^2)"
    return [area_pixels, area_microns_squared]
}

// Function to get a rectangle annotation of a fixed size around a given cell detection object
def get_window_around_cell(cell_roi, width, height) {
    // Compute center coordinates of cell detection object
    x = cell_roi.getCentroidX() - width/2
    y = cell_roi.getCentroidY() - height/2

    // Create a rectangle window around the cell and add it to the annotation objects. Return this annotation
    rectangleROI = ROIs.createRectangleROI(x, y, width, height, ImagePlane.getDefaultPlane())
    window_annotation = PathObjects.createAnnotationObject(rectangleROI)
    addObjects(window_annotation)
    return window_annotation
}

// Function to check if the overlap between a new window and the previous windows exceeds a certain threshold.
// If this is the case, it means that the new detected cell inside the window is likely the same cell as another, already processed one.
def check_if_new_cell(all_cell_windows, new_window, threshold) {
    max_overlap_area = (new_window.getROI().x2 - new_window.getROI().x) * (new_window.getROI().y2 - new_window.getROI().y)
    if (all_cell_windows.isEmpty()) {
        return true
    }
    for (window in all_cell_windows) { 
        x_overlap = Math.max(0, Math.min(window.getROI().x2, new_window.getROI().x2) - Math.max(window.getROI().x, new_window.getROI().x))
        y_overlap = Math.max(0, Math.min(window.getROI().y2, new_window.getROI().y2) - Math.max(window.getROI().y, new_window.getROI().y))
        overlap_area = x_overlap * y_overlap
        overlap_ratio = overlap_area / max_overlap_area
        //println "X overlap: ${x_overlap}, Y overlap: ${y_overlap}, Overlapping area: ${overlap_area}, max area: ${max_overlap_area} -> Ratio: ${overlap_ratio}"
   
        if (overlap_ratio >= threshold) {
            println "Overlapping cell found - do not process this cell"
            return false
        }
    }
    return true
}


// ===================================== STAIN VECTOR PREPROCESSING =============================================

// Function to manually fix and set the stain vectors (swaps DAB & Hematoxylin channels)
def set_stain_vectors_swapped() {
    println "Original stain vector values: ${imageData.getColorDeconvolutionStains()}"
    setImageType('BRIGHTFIELD_H_DAB');
    setColorDeconvolutionStains("""{"Name" : "Brightfield (H-DAB)",
                                "Stain 1" : "DAB", "Values 1" : "0.58 0.732 0.357", 
                                "Stain 2" : "Hematoxylin", "Values 2" : "0.276 0.546 0.791",
                                "Background" : " 234 230 226"}""")
    println "Set new stain vector values, swapping DAB & Hematoxylin: ${imageData.getColorDeconvolutionStains()}"
    //println "Stain vectors: Stain 1: ${imageData.getColorDeconvolutionStains().stain1}, Stain 2: ${imageData.getColorDeconvolutionStains().stain2}"
}

// Function to check if the stain vectors for the hematoxylin & DAB channel have already been swapped
def check_if_swapped() {
    stain_vectors = imageData.getColorDeconvolutionStains()
    vector_1_red = stain_vectors.getStain(1).getRed()
    vector_2_red = stain_vectors.getStain(2).getRed()
    //println "Stain vectors: Stain 1: ${stain_vectors.stain1}, Stain 2: ${stain_vectors.stain2}"
    //println "RED1 ${vector_1_red}, RED2 ${vector_2_red} - ${vector_1_red < vector_2_red} - ALREADY SWAPPED"
    if (vector_1_red < vector_2_red) {
        println "Stain vectors were already swapped before. Skip swapping"
        return true
    }
    return false
}

// Function to swap the stain vectors hematoxylin & DAB
def swap_stain_vectors() {
    println "Swapping stain vectors: DAB & Hematoxylin"
    // Get estimated stain vectors
    stain_vectors = imageData.getColorDeconvolutionStains()
    hematoxylin_vector = stain_vectors.getStain(1)
    dab_vector = stain_vectors.getStain(2)
 
    // Swap the stain vectors
    stain_vectors.stain2 = hematoxylin_vector
    stain_vectors.stain1 = dab_vector
    //println "Stain vectors: Stain 1: ${stain_vectors.stain1}, Stain 2: ${stain_vectors.stain2}"
    imageData.setColorDeconvolutionStains(stain_vectors)
}


// ===================================== SOMA CELL DETECTION & FILTERING ===================================

// Function to run the cell detection algorithm to find soma body cells in the DAB-channel
def detect_soma_body_cells(min_area, threshold, circularity, min_caliper, mean_hematoxylin) {
    println "Starting cell detection"
    runPlugin('qupath.imagej.detect.cells.WatershedCellDetection', 
              '{"detectionImageBrightfield":"Hematoxylin OD","requestedPixelSizeMicrons":0.5,' +
              '"backgroundRadiusMicrons":8.0,"backgroundByReconstruction":true,"medianRadiusMicrons":0.0,'+
              '"sigmaMicrons":1.5,"minAreaMicrons":'+ min_area +',"maxAreaMicrons":400.0,"threshold":'+ threshold +
              ',"maxBackground":2.0,"watershedPostProcess":true,"excludeDAB":false,"cellExpansionMicrons":0.0,' +
              '"includeNuclei":true,"smoothBoundaries":true,"makeMeasurements":true}')
    
    // Postprocess cell detection: Filter out cells
    filter_out_cells(circularity, min_caliper, mean_hematoxylin)
}

// Function to filter out cell detections based on multiple criteria (cell circularity, min caliper, mean hematoxylin (dab) stain)
def filter_out_cells(circularity_threshold, min_caliper_threshold, mean_hematoxylin_threshold) {
    removed_cells = 0
    for (detection in getDetectionObjects()) {
        double circularity = measurement(detection, "Nucleus: Circularity")
        double caliper = measurement(detection, "Nucleus: Min caliper")
        double mean_hematoxylin_intensity = measurement(detection, "Nucleus: Hematoxylin OD mean")

        if (circularity < circularity_threshold || caliper < min_caliper_threshold || mean_hematoxylin_intensity < mean_hematoxylin_threshold) {
           removeObject(detection, true)
           removed_cells++
        }
    }
    println "Filtering cells: Removed ${removed_cells} cells that didn't meet the specified thresholds for minimum circularity, min caliper & mean hematoxylin intensity"
}


// ===================================== IMAGE-J MAKRO INTEGRATION ==========================================

// Function to initialize the ImageJ macro runner (what is sent to & retrieved from ImageJ, the downsampling rate etc.)
def instantiate_imageJ_macro(downsample_factor) {
    println "Instantiating ImageJ Macro Runner"
    params = new ImageJMacroRunner(getQuPath()).getParameterList()
    params.getParameters().get('downsampleFactor').setValue(downsample_factor)
    params.getParameters().get('sendROI').setValue(true)
    params.getParameters().get('sendOverlay').setValue(false)
    params.getParameters().get('doParallel').setValue(false)
    params.getParameters().get('clearObjects').setValue(false)
    params.getParameters().get('getROI').setValue(true)
    params.getParameters().get('getOverlay').setValue(false)
    params.getParameters().getOverlayAs.setValue("Detections")
    return params
}

// Function to create the string that implements the ImageJ macro used for preprocessing, skeletonization & quantification of a soma cell
def skeletonization_analysis_macro(save_path, cell_counter, gaussian_sigma, bp_filter_large, bp_filter_small, floodfill_pos, \
                                   threshold_variant, global_threshold_type, local_threshold_type, local_threshold_radius) {
   //println "Loading ImageJ macro script" 
   error = false
   
   // Get all individual macro scripts & concatenate them to a full script
   if (threshold_variant == "global") {
       preprocessing_macro = get_preprocesing_macro_globalTR(bp_filter_large, bp_filter_small, gaussian_sigma, global_threshold_type)
   } else if (threshold_variant == "local") {
       preprocessing_macro = get_preprocessing_macro_localTR(bp_filter_large, bp_filter_small, gaussian_sigma, local_threshold_type, local_threshold_radius)
   } else {
       preprocessing_macro = ""
       error = true
   }
   binary_mask_macro = get_binary_mask_macro(floodfill_pos)
   skeletonization_macro = get_skeletonization_quantification_macro(save_path, cell_counter)
 
   full_macro = preprocessing_macro + binary_mask_macro + skeletonization_macro
   //println "MACRO: ${full_macro}"
   return [full_macro, error]
}

// Preprocessing Macro Variant 1: Color Deconvolution, Bandpass-Filter, Gaussian Blur & Auto Global Thresholding
def get_preprocesing_macro_globalTR(bp_filter_large, bp_filter_small, gaussian_sigma, global_threshold_type) {
    // In case we want to use unsharp mask again, insert this line (after invert & before gaussian blur):
    // 'run("Unsharp Mask...", "radius=' + unsharp_mask_r + ' mask=' + unsharp_mask_m + '");' + 'run("Despeckle");' +
    preprocessing_macro = 'org_img_id = getImageID();' +
                          'selectImage(org_img_id);' +
                          'imageTitle = getTitle();' +
                          'run("Colour Deconvolution", "vectors=[H DAB] show");' +
                          'selectWindow(imageTitle + "-(Colour_2)");' +
                          'run("Bandpass Filter...", "filter_large=' + bp_filter_large + ' filter_small=' + bp_filter_small + ' suppress=None tolerance=5 autoscale saturate");' +
                          'run("8-bit");' +
                          'run("Invert");' +
                          'run("Gaussian Blur...", "sigma=' + gaussian_sigma + '");' +
                          'run("Auto Threshold", "method=' + global_threshold_type + ' white");' +
                          'run("Fill Holes");'
    return preprocessing_macro
}

// Preprocessing Macro Variant 2: Color Deconvolution, Bandpass-Filter, Gaussian Blur & Auto Local Thresholding
def get_preprocessing_macro_localTR(bp_filter_large, bp_filter_small, gaussian_sigma, local_threshold_type, local_threshold_radius) {
    // In case we want to use unsharp mask again, insert this line (after invert & before gaussian blur):
    // 'run("Unsharp Mask...", "radius=' + unsharp_mask_r + ' mask=' + unsharp_mask_m + '");' + 'run("Despeckle");' +
    preprocessing_macro = 'org_img_id = getImageID();' +
                          'selectImage(org_img_id);' +
                          'imageTitle = getTitle();' +
                          'run("Colour Deconvolution", "vectors=[H DAB] show");' +
                          'selectWindow(imageTitle + "-(Colour_2)");' +
                          'run("Bandpass Filter...", "filter_large=' + bp_filter_large + ' filter_small=' + bp_filter_small + ' suppress=None tolerance=5 autoscale saturate");' +
                          'run("8-bit");' +
                          'run("Invert");' +
                          'run("Gaussian Blur...", "sigma=' + gaussian_sigma + '");' +
                          'run("Auto Local Threshold", "method=' + local_threshold_type + ' radius=' + local_threshold_radius + ' parameter_1=0 parameter_2=0 white");' +
                          'run("Fill Holes");'
    return preprocessing_macro
}

// Macro script that creates a new binary mask showing only the central soma body cell & its branches
def get_binary_mask_macro(floodfill_pos) {
    binary_mask_macro = 'ID_mask = getImageID();' +
                        'run("Duplicate...", " ");' +
                        'duplicated_mask_id = getImageID();' +
                        'selectImage(duplicated_mask_id);' +
                        'run("Colors...", "foreground=black background=white selection=yellow");' +
                        'floodFill(' + floodfill_pos + ', ' + floodfill_pos + ', 0);' +
                        'ID_mask_filled = getImageID();' +
                        'imageCalculator("Subtract create", ID_mask, ID_mask_filled);'
    return binary_mask_macro
}

// Macro script that runs skeletonization & saves the quantified branch-results to a csv-file. It also returns the skeleton as overlay to QuPath
def get_skeletonization_quantification_macro(save_path, cell_counter) {
    skeletonization_macro = 'ID_mircoglia_mask = getImageID();' +
                            'selectImage(ID_mircoglia_mask);' +
                            'run("Analyze Particles...", "size=1-Infinity add");' +
                            'run("Skeletonize");' +
                            'run("8-bit");' +
                            'run("Properties...", "voxel_depth=1");' +
                            'run("Analyze Skeleton (2D/3D)", "prune=[shortest branch] show display");' +
                            'selectWindow("Results");' +
                            'saveAs("Results", "' + save_path + '/results_cell_' + cell_counter + '.csv");' +
                            'selectImage(ID_mircoglia_mask);' +
                            'run("Create Selection");' +
                            'selectImage(org_img_id);' +
                            'run("Restore Selection");'
    return skeletonization_macro
}
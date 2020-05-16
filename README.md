# pytesseract_utils
Some handy features for pytesseract OCR


* A set of features for OCR-Detection with pytesseract.

Features:
- OCR from file (with table content) and save result per file as Excel-file (Multithreaded).
- OCR from file and save result per file as textfile or save results in one textfile (Multithreaded).
- OCR from screenshot, copy result as plain text without linebreaks, as rows (with linebreaks) or columns to clipboard.
- OCR from screenshot of table, save result as table.

#######################################################################
BEFORE YOU START:
#######################################################################
Install tesseract from https://github.com/UB-Mannheim/tesseract/wiki
Install requirements

#######################################################################
USAGE:
#######################################################################
REPLACE >>r'...\tesseract.exe'<< with your path to the tesseract.exe

#######################################################################
# OCR FROM SCREENSHOT TO CLIPBOARD
#######################################################################
# draw a border around table-content without border
prep = prepare_draw_table(search_vertical = False)
output_image = prep.run("input_image.jpg") 

ocr2c = screenshotOCR_to_clipboard(r'...\tesseract.exe')
result_as_line = ocr2c.ocr_to_oneline()
print("Result oneline:", result_as_line)

result_as_rows = ocr2c.ocr_to_rows()
print("Result rows:", result_as_rows)

result_as_cols = ocr2c.ocr_to_cols()
print("Result cols:", result_as_cols)
ocr2c.copy_to_clipboard()

result_as_df = ocr2c.ocr_table_to_table()
print("result_as_df:", result_as_df)

ocr2c.copy_to_clipboard()
cv2.destroyAllWindows()


#######################################################################
# OCR FROM FILES
#######################################################################
start =  time.time()
# Set your options here
# WHERE ARE THE FILES
input_dir = "input"
# WHERE TO STORE THE RESULTS
output_dir = "ocr_results"
# SET THE NUMBER OF THREADS
number_of_threads = multiprocessing.cpu_count()
tesseract_pth = r'...\tesseract.exe'
ocr = fileOCR_table_to_xlsx(tesseract_pth, input_dir, output_dir, cpu_threads = number_of_threads, debug = FALSE)
ocr.run()
print("DONE!\nThis took: {} seconds".format(time.time()-start))

start =  time.time()
# Set your options here
# WHERE ARE THE FILES
input_dir = "input"
# WHERE TO STORE THE RESULTS
output_dir = "ocr_results"
# SET THE NUMBER OF THREADS
number_of_threads = multiprocessing.cpu_count()
tesseract_pth = r'...\tesseract.exe'
ocr = fileOCR_text_to_textfile(tesseract_pth, input_dir, output_dir, cpu_threads = number_of_threads, debug = True)
# SHOW HELP
ocr.help()
# SEE HELP TO GET A HINT FOR USAGE OF all_to_one_file AND resize_faktor
ocr.run(all_to_one_file = True, resize_faktor = 2)
print("DONE!\nThis took: {} seconds".format(time.time()-start))


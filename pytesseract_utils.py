# -*- coding: utf-8 -*-
"""
Created on Mon May  4 2020

@author: Lutz Kaiser
"""
import time
from PIL import ImageGrab
import cv2
import numpy as np
import win32gui, win32con, win32com.client
#import pymouse
import ctypes
from ctypes import windll, Structure, c_long, byref
user32 = ctypes.windll.user32
screensize = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)

import pytesseract
import pyperclip
import re
import pandas as pd

import os
import pdf2image
import matplotlib.pyplot as plt
import multiprocessing
multiprocessing.freeze_support()
import threading
try:
    from PIL import Image
except ImportError:
    import Image

# To get the PID of a process
import psutil

# =============================================================================
# 
# =============================================================================
class screenshot:
    def __init__(self,):
        self.refPt = []
        self.x = 0
        self.y = 0
        self.img = None
        self.img_orig = None
        
        # multiple selection from screenshot
        self.rois = []
        
        self.result = ""

    def clickposition(self, event, x,y, flags, param):
        # print("[clickposition]")
        if event == cv2.EVENT_LBUTTONDOWN:
            self.refPt = [(x, y)]
            # print("Start x, y", x,y)
        elif event == cv2.EVENT_LBUTTONUP:
            self.refPt.append((x,y))
            # Draw selected area
            self.img = cv2.rectangle(self.img, self.refPt[0], (x,y), (255,0,0))
            # save the selected roi from picture without roi-borders
            self.rois.append(self.crop(self.img_orig))
            # print("End x, y", x,y)
        elif len(self.refPt) == 1:
            # Draw rectangle while selecting
            img_copy = self.img.copy()
            img_copy = cv2.rectangle(img_copy, self.refPt[0], (self.x, self.y ), (255,0,0))
            cv2.imshow("Screenshot", img_copy)
        self.x = x
        self.y = y
        
    def crop(self, img):
        # creating the roi
        start, end = self.refPt
        roi = img[start[1]:end[1],start[0]:end[0]]
        return roi
    
    def make_screenshot(self):
        # Screenshot des gesamten Bildschirms
        screen_width =  screensize[0]
        screen_height = screensize[1]
        
        area=(0, 0, screen_width, screen_height)
        
        x1 = min(int(area[0]), screen_width)
        y1 = min(int(area[1]), screen_height)
        x2 = min(int(area[2]), screen_width)
        y2 = min(int(area[3]), screen_height)
        
        search_area = (x1, y1, x2, y2)
        
        img = ImageGrab.grab((0,0, screen_width, screen_height)) # .convert("BGRA")
        img_cv = np.array(img)
        img_rgb = img_cv[:, :, ::-1].copy()
        return img_rgb
    
    def select_roi(self, img_rgb):       
        title = "Screenshot"
        cv2.namedWindow(title)
        cv2.setMouseCallback(title, self.clickposition)
        cv2.imshow(title, img_rgb)
        # maximize the window
        self.max_scrnshot_window(title)
        cv2.waitKey(0)
    
    def minimize_window(self):
        # Minimize current Window
        Minimize = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(Minimize, win32con.SW_MINIMIZE)
        time.sleep(0.1)
        
    # func to get the screenshot-window
    def max_scrnshot_window(self, p_name = ""):
        # Takes a window-name, searches for the corresponding PID and maximize the window    
        p_pid = win32gui.FindWindow(None, p_name)
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(p_pid)
        win32gui.ShowWindow(p_pid, win32con.SW_MAXIMIZE)
        
    def run(self, minimize_window = False):
        # Takes a screenshot of the Desktop
        if minimize_window:
            self.minimize_window()
        self.img = self.make_screenshot()
        # copy picture for a "clean" version
        self.img_orig = self.img.copy()
        # Shows the screenshot. User has to select a roi to crop to
        self.select_roi(self.img)
        # get the last selected roi
        for roi in self.rois:
            # roi = self.crop(self.img)
            cv2.imshow("ROI", roi)
            cv2.waitKey(0)
        print("Screenshots finished")
        cv2.destroyAllWindows()
        return self.rois


class prepare_draw_table:
    """
    MAKES A TABLE OF NO-TABLE CONTENT FOR OUTPUT AS AN EXCEL-FILE
    """
    def __init__(self, line_space = 5, line_offset = 3, hor_space = 15, line_thickness = 2, search_vertical = False):
        # Horizontale Linien
        self.h_lines = []
        # Zeilenabstand
        self.space = line_space
        
        # Vertikale Linien
        self.v_lines = []
        # Horizontaler Schriftabstand
        self.hor_space = hor_space
        
        # Konsekutive Pixel mit Zeilenabstand
        self.found_space = True
        # Abstand unter der Schrift bei gefundenem Zeilenabstand
        self.line_offset = line_offset
        # Abstand zum Bildrand
        self.offset_outerborder = 10
        
        # Threshold, bei dem keine Schrift vermutet wird => Möglichst klein, wegen
        # zeilenübergreifenden Buchstaben (j,g,p,q)
        self.ls_thresh = 10
        
        # Linenstärke
        self.line_thick = line_thickness
        
        # vertikale Suche
        self.search_vertical = search_vertical
        
    def search_for_horlines(self, image):
        # Untersucht das Bild von oben nach unten auf Zwischenräume 
        img_height = image.shape[0]
        # Bildtransformation SW
        img_copy = image.copy()
        # Bild Invertieren
        img_copy = 255-img_copy

        cv2.imshow("Copy", img_copy)
        cv2.waitKey(0)
        
        thresh, img_sw = cv2.threshold(img_copy, 64, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        # Über die Bildhöhe iterieren
        for i in range(img_height - self.space):
            # Auf Zeilenabstand prüfen
            line_space = img_sw[i:i+self.space, 0::]
            # Count Nonzeros
            nzs = np.count_nonzero(line_space)
            # Schrift ist weiß
            # nzs < Thresh => keine Schrift
            if nzs <= self.ls_thresh and not self.found_space:
                # Speicher die Position des Zwischenraums
                self.h_lines.append(i)
                self.found_space = True
            elif nzs >= self.ls_thresh and self.found_space:
                # setzt den Pixelzähler zurück
                self.found_space = False
                
                
    def search_for_verlines(self, image):
        # self.found_space = True
        # Untersucht das Bild von links nach rechts auf Zwischenräume 
        # ==> Linien von oben nach unten
        img_width = image.shape[1]
        # Bildtransformation SW
        img_copy = image.copy()
        img_copy = 255-img_copy
        # cv2.imshow("Copy", img_copy)
        thresh, img_sw = cv2.threshold(img_copy, 64, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        # Über die Bildbreite iterieren
        for i in range(img_width - self.hor_space):
            # Auf Zeilenabstand prüfen
            line_space = img_sw[0::, i:i+self.hor_space]
            # Count Nonzeros
            nzs = np.count_nonzero(line_space)
            # nzs > Thresh => Schrift, sonst space
            if nzs <= self.ls_thresh and not self.found_space:
                # Speicher die Position des Zwischenraums
                self.v_lines.append(i)
                self.found_space = True
            elif nzs >= self.ls_thresh and self.found_space:
                # setzt den Pixelzähler zurück
                self.found_space = False
                
    def draw_horlines(self, image):
        # Zeichnet die Linien aus search_for_horlines in ein Bild
        img_bordered = image.copy()
        w = image.shape[1]
        offset = self.offset_outerborder
        for p in self.h_lines:
            img_bordered = cv2.line(img_bordered, (0+offset, p+self.line_offset), (w -1 -offset, p+self.line_offset), 0, self.line_thick )
        # cv2.imshow("HLines", img_bordered)
        # cv2.waitKey(0)
        return img_bordered
        
    
    def draw_verlines(self, image):
        # Zeichnet die Linien aus search_for_verlines in ein Bild
        img_bordered = image.copy()
        h = image.shape[0]
        offset = self.offset_outerborder
        for p in self.v_lines:
            # p(x,y)
            img_bordered = cv2.line(img_bordered, (p+int(self.hor_space/2), 0+offset), (p+int(self.hor_space/2), h -1 -offset), 0, self.line_thick )
        cv2.imshow("VLines", img_bordered)
        cv2.waitKey(0)
        return img_bordered
    
    def draw_outer_borders(self, image):
        # Zeichnet Rahmen um das gesamte Bild
        h, w = image.shape
        offset = 10
        image = cv2.rectangle(image, (0+offset,0+offset), (w-1-offset,h-1-offset), 0, self.line_thick )
        cv2.imshow("Outb", image)
        cv2.waitKey(0)
        return image
        
    def run(self, img_pth):
        # Führt die einzelnen Funktionen konsekutiv aus
        image = cv2.imread(img_pth)
        image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        self.search_for_horlines(image)
        img_lines = self.draw_horlines(image)
        cv2.imshow("HLines", img_lines)
        cv2.waitKey(0)
        if self.search_vertical:
            self.search_for_verlines(image)
            img_lines = self.draw_verlines(img_lines)
            # cv2.imshow("VLines", img_lines)
            # cv2.waitKey(0)
        img_f = self.draw_outer_borders(img_lines)
        return img_f


class screenshotOCR_to_clipboard:
    # Makes a screenshot, pass the image to ocr and copy result to clipboard
    
    def __init__(self, tesseract_pth, debug = False):
        self.scrnshot = screenshot()
        # Path to .exe of tesseract
        pytesseract.pytesseract.tesseract_cmd = tesseract_pth
        self.debug = debug
    
    def get_screenshot(self, minimize_window = False):
        rois = self.scrnshot.run(minimize_window)
        return rois
    
    def img_preprocessing(self, roi):
        roi = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        img_copy = 255-roi
        # if self.debug:
            # cv2.imshow("Copy", img_copy)
            # cv2.waitKey(0)        
        
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 1))
        # border = cv2.copyMakeBorder(img_sw, 2,2,2,2, cv2.BORDER_CONSTANT,value=[255,255])
        resizing = cv2.resize(img_copy, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        thresh, resizing = cv2.threshold(resizing, 64, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        dilation = cv2.dilate(resizing, kernel,iterations=1)
        erosion = cv2.erode(dilation, kernel,iterations=2)
        return erosion
        
    def get_ocr_result(self, roi):
        roi_cnvrt = self.img_preprocessing(roi)
        if self.debug:
            cv2.imshow("roi_cnvrt", roi_cnvrt)
            cv2.waitKey(0)
        return pytesseract.image_to_string(roi_cnvrt)
    
    def text_postprocessing(self, text):
        # get rid off linebreaks
        text = text.replace("\r\n"," ").replace("\n"," ").replace("  ", " ")
        return text
    
    def copy_to_clipboard(self):     
        # cmd = 'echo ' + self.result.replace("  ", " ").strip() + '| clip'  # echo | set /p nul=
        # os.system(cmd)
        if self.debug:
            print("copying:", self.result.replace("  ", " ") )
        pyperclip.copy(self.result.replace("  ", " "))
    
    def ocr_to_oneline(self):
        # Results will have no linebreaks => result in one line
        rois = self.get_screenshot(minimize_window = True)
        results = []
        for roi in rois:
            result = self.get_ocr_result(roi)
            results.append(result)
        # print(results)
        res = ""
        for result in results:
            res += self.text_postprocessing(result) + " "
        # print(res)
        self.result = res
        return res
        
    def ocr_to_rows(self):
        # Results will have linebreaks
        rois = self.get_screenshot(minimize_window = True)
        results = []
        for roi in rois:
            result = self.get_ocr_result(roi)
            results.append(result)
        # print(results)
        res = ""
        for result in results:
            res += result + " "
        # print(res)
        self.result = res
        return res
    
    def ocr_to_cols(self, seperator = "\t"):
        # Results will have seperators for saving as excel-file
        # check excel settings to get the needed seperator
        rois = self.get_screenshot(minimize_window = True)
        results = []
        for roi in rois:
            result = self.get_ocr_result(roi)
            results.append(result)

        res = ""
        for result in results:
            # Remove empty rows
            if result not in [" ", "", "\n", "\r\n"]:
                res += result + " "
        # search for linebreaks
        if re.search("\r\n", res) is not None:
            res = res.replace("\r\n", seperator)
        else:
            res = res.replace("\n", seperator)
        # remove double seperators
        res = res.replace(seperator+seperator,seperator)
        self.result = res
        return res
    
    def ocr_table_to_table(self, rows_to_cols = False):
        # Make a screenshot
        # Pass it to image processing with cv2 and get the cell boxes
        # Pass the boxes to tesseract
        # make a excel dataframe of the results
        # copy dataframe to clipboard
        rois = self.get_screenshot(minimize_window = True)
        for roi in rois:
            # ToDo: How to deal with multiple rois
            finalboxes, countcol, row, bitnot_img = self.do_image_processing(roi)
            result = self.get_ocrtext(finalboxes, bitnot_img)
            dataframe = self.create_dataframe(result, countcol, row, rows_to_cols = rows_to_cols)
            # copy to clipboard
            dataframe.to_clipboard()
            print("Ready to paste")
            # Give the user time to paste
            time.sleep(2)
  
    def do_image_processing(self, screenshot, invert_image = True, max_cell_width = 1000, max_cell_height = 500):      
        
        # Image aus screenshot erhalten
        img = screenshot
         
        # count the loops
        pcnt = 0
        img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        while True:
            # Pass the screenshot to preprocessing
            # img = self.img_preprocessing(img)
            
            # Convert to binary
            thresh, img_bin = cv2.threshold(img, 64, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
            cv2.imshow("img_bin", img_bin)
            cv2.waitKey(0)
            
            #inverting the image
            if invert_image:
                img_bin = 255-img_bin
            
            # countcol(width) of kernel as 100th of total width
            kernel_len = np.array(img).shape[1]//100
            # Defining a vertical kernel to detect all vertical lines of image 
            ver_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, kernel_len))
            # Defining a horizontal kernel to detect all horizontal lines of image
            hor_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_len, 1))
            # A kernel of 2x2
            kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
            
            #Use vertical kernel to detect and save the vertical lines in a jpg
            image_1 = cv2.erode(img_bin, ver_kernel, iterations=3)
            vertical_lines = cv2.dilate(image_1, ver_kernel, iterations=3)
            
            #Use horizontal kernel to detect and save the horizontal lines in a jpg
            image_2 = cv2.erode(img_bin, hor_kernel, iterations=3)
            horizontal_lines = cv2.dilate(image_2, hor_kernel, iterations=3)
            # imgname = self.check_for_dublicate("{}_horizontal".format(fname), self.out_pth + "/images", ext = ".jpg" )
            # cv2.imwrite(self.out_pth + "/images/" + imgname + ".jpg", horizontal_lines)
            
            # Combine horizontal and vertical lines in a new third image, with both having same weight.
            img_vh = cv2.addWeighted(vertical_lines, 0.5, horizontal_lines, 0.5, 0.0)
            #Eroding and thesholding the image
            img_vh = cv2.erode(~img_vh, kernel, iterations=2)
            thresh, img_vh = cv2.threshold(img_vh, 128,255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
            # imgname = self.check_for_dublicate("{}_img_vh".format(fname), self.out_pth + "/images", ext = ".jpg" )
            # cv2.imwrite(self.out_pth + "/images/" + imgname + ".jpg", img_vh)
            cv2.imshow("img_vh", img_vh)
            cv2.waitKey(0)
            # combining borders and text
            bitxor = cv2.bitwise_xor(img, img_vh)
            bitnot = cv2.bitwise_not(bitxor)
            bitnot_img = bitnot
            
            # Detect contours for following box detection
            contours, hierarchy = cv2.findContours(img_vh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
            
            def sort_contours(cnts, method="left-to-right"):
                # initialize the reverse flag and sort index
                reverse = False
                i = 0
                # handle if we need to sort in reverse
                if method == "right-to-left" or method == "bottom-to-top":
                    reverse = True
                # handle if we are sorting against the y-coordinate rather than
                # the x-coordinate of the bounding box
                if method == "top-to-bottom" or method == "bottom-to-top":
                    i = 1
                # construct the list of bounding boxes and sort them from top to
                # bottom
                boundingBoxes = [cv2.boundingRect(c) for c in cnts]
                (cnts, boundingBoxes) = zip(*sorted(zip(cnts, boundingBoxes),
                key=lambda b:b[1][i], reverse=reverse))
                # return the list of sorted contours and bounding boxes
                return (cnts, boundingBoxes)
            
            # Sort all the contours by top to bottom.
            contours, boundingBoxes = sort_contours(contours, method="top-to-bottom")
            
     
            # Create list box to store all boxes in  
            box = []
            heights = []
            
            # Get position (x,y), width and height for every contour and show the contour on image
            for c in contours:
                x, y, w, h = cv2.boundingRect(c)
                print("w", w, "h", h)
                
                # Limit the size of a cell
                if (w < max_cell_width and h < max_cell_height):
                    image = cv2.rectangle(img,(x,y),(x+w,y+h),(0,255,0), 4)
                    box.append([x,y,w,h])
                    heights.append(h)
            
            #Creating two lists to define row and column in which cell is located
            row=[]
            column=[]
            j = 0
            i = 0
            
            for i in range(len(box)):    
                # box = [[x, y, w, h], ]
                # iterating over the number of rois found by cv2.cont
                # take care of change in rows and cols
                if(i==0):
                    column.append(box[i])
                    previous=box[i]    
                
                else:
                    # checking for changed row position
                    # opencv is going from 0 up
                    if(box[i][1] <= previous[1] + previous[3]/2):  # mean_height/2
                        column.append(box[i])
                        previous = box[i]            
                        # add to row if end box-list is reached
                        if(i == len(box) -1):
                            row.append(column)        
     
                    else:                        
                        row.append(column)
                        column=[]
                        previous = box[i]
                        column.append(box[i])
                        
            if self.debug:
                print("\nCOL:\n", len(column), column)
                print("\nROW:\n", len(row), row)
                # input()
            #calculating maximum number of cells
            countcol = 0
            for i in range(len(row)):
                countcol = len(row[i])
                if countcol > countcol:
                    countcol = countcol

            
            center = []
            if countcol != 0 and len(row) != 0:
                #Retrieving the center of each column
                center = [int(row[i][j][0]+row[i][j][2]/2) for j in range(len(row[i])) if row[0]]
                center=np.array(center)
                center.sort()
            if self.debug:
                print("\nCENTER:\n", center)
                
            #Regarding the distance to the columns center, the boxes are arranged in respective order
            finalboxes = []
            for i in range(len(row)):
                lis=[]
                for k in range(countcol):
                    lis.append([])
                for j in range(len(row[i])):
                    diff = abs(center-(row[i][j][0]+row[i][j][2]/4))
                    minimum = min(diff)
                    indexing = list(diff).index(minimum)
                    lis[indexing].append(row[i][j])
                finalboxes.append(lis)
                
            # leave the loop if table border was found
            if len(finalboxes) != 0:
                break
            else:
                # break
                if pcnt == 1:
                    break
                print("[INFO] Got no borders. Trying to sharpen the image")
                # # do some image transformation
                # try to sharpen the picture
                kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]])
                # img = cv2.filter2D(img, -1, kernel)
                img = cv2.filter2D(img, -1, kernel)
                img = cv2.GaussianBlur(img, (3,3),0)
                pcnt += 1
                
                
        return finalboxes, countcol, row, bitnot_img
            
        

    def get_ocrtext(self, finalboxes, bitnot_img):
        print("[INFO] Getting OCR Text" )
        #from every single image-based cell/box the strings are extracted via pytesseract and stored in a list
        outer=[]
        for i in range(len(finalboxes)):
            for j in range(len(finalboxes[i])):
                inner=''
                if(len(finalboxes[i][j])==0):
                    outer.append(' ')
                else:
                    for k in range(len(finalboxes[i][j])):
                        y,x,w,h = finalboxes[i][j][k][0],finalboxes[i][j][k][1], finalboxes[i][j][k][2],finalboxes[i][j][k][3]
                        finalimg = bitnot_img[x:x+h, y:y+w]
                        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 1))
                        border = cv2.copyMakeBorder(finalimg,2,2,2,2, cv2.BORDER_CONSTANT,value=[255,255])
                        resizing = cv2.resize(border, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
                        dilation = cv2.dilate(resizing, kernel,iterations=1)
                        erosion = cv2.erode(dilation, kernel,iterations=2)
                        # cv2.imshow("Erosion", erosion)
                        # cv2.waitKey(0)
                        out = pytesseract.image_to_string(erosion)
                        if(len(out)==0):
                            # If the rusult was zero, try different settings
                            out = pytesseract.image_to_string(erosion, config='--psm 3')
                        inner = inner +" "+ out
                        # inner = self.text_postproc(inner)
                    
                    if self.debug:
                        print(inner)
                    outer.append(inner)
        return outer       
          
    def create_dataframe(self, ocr_data, countcol, row, rows_to_cols):      
        print("[INFO] Creating Dataframe" )
        #Creating a dataframe of the generated OCR list
        arr = np.array(ocr_data)
        if not rows_to_cols:
            # Cols are cols, rows are rows
            dataframe = pd.DataFrame(arr.reshape(len(row), countcol))
        else:
            # Rows become cols, cols become rows
            dataframe = pd.DataFrame(arr.reshape(countcol, len(row)))
        print("\nDATAFRAME:\n", dataframe)
        return dataframe
        


class fileOCR_table_to_xlsx:
    """
    From file with table content to .xlsx file
    """
    def __init__(self, tesseract_pth, input_path = "", output_path = "", use_all_files = True, show_pics = False, cpu_threads = 4, 
                 convrt2table = True, debug = True):
        self.data_dir = input_path if not input_path == "" else "input"
        self.out_pth = output_path if not output_path == "" else "results"
        os.makedirs(self.out_pth, exist_ok=True)
        self.all_fs = use_all_files
        self.data = []
        self.show_pics = show_pics
        self.debug = debug
        if self.debug:
            os.makedirs(self.out_pth + "/images", exist_ok=True)
        self.finalboxes = []
        self.bitnot_img = None
        
        # if pdf file has more than one site, the converted images will be stored here
        self.cv_images = []
        
        # Multithreading
        self.cpu_threads = cpu_threads
        
        # save the pdf_name or gif name
        self.converted_file = ""
        
        # Falls ein Bild in eine Tabellenform gebracht werden soll
        self.convrt2table = convrt2table
        self.mop = prepare_draw_table(search_vertical = False)
        
        # Transform result from cols to rows
        # self.rows_to_cols = False
        
        # Check tesseract path
        if os.path.isfile(tesseract_pth):
            pytesseract.pytesseract.tesseract_cmd = tesseract_pth
        else:
            print("[ERROR] No valid tesseract.exe")
            raise
        
    def load_data(self):
        raw_data = os.listdir(self.data_dir)
        if not self.all_fs:
            # valid datatypes: jpg, jpeg, png, pdf
            vext = ["jpg", "jpeg", "png", "pdf"]
            
            for file in raw_data:
                ext = file.split(".")[-1]
                if ext in vext:
                    file = self.data_dir + "/" + file
                    self.data.append(file)
        else:
            for file in raw_data:
                file = self.data_dir + "/" + file
                self.data.append(file)
                
    def convrt_pdf(self, pdf_file):
        name = pdf_file.split("/")[-1].replace(".","_")
        self.converted_file = name
        pilimages = pdf2image.convert_from_path(pdf_file)
        cnt = 0
        for pilimage in pilimages:
            cv_image = np.array(pilimage) 
            # Convert RGB to BGR 
            cv_image = cv_image[:, :, ::-1]
            if self.debug:
                cv2.imwrite("temp/img_from_pdf_" + str(cnt)+".jpg", cv_image)
                cnt += 1
            self.cv_images.append(cv_image)
            
    
    def do_image_processing(self, file, fname, file_is_path = True):      
        
        if file_is_path:
            if type(file) == str:            
                img = cv2.imread(file, 0)
                # print(type(img))
            else:
                # got a converted pdf image
                # print(type(file))
                os.makedirs("temp", exist_ok=True)
                temp_name = "temp_{}".format(fname)
                fname = self.check_for_dublicate(temp_name, "temp", ext = ".png")
                cv2.imwrite("temp/{}.png".format(fname), file)
                img = cv2.imread("temp/{}.png".format(fname), 0)
                # try:
                #     os.remove("temp/{}.png".format(fname))
                # except Exception:
                #     pass
        else:
            # Image aus ocr_preps erhalten
            img = file
           
        # backup the img
        orig_img = img
        # count the loops
        pcnt = 0
        
        # for binary image
        lower_thresh = 64
        upper_thresh = 255
        
        
        borderless_img = None
        
        while True:
            
            #thresholding the image to a binary image
            # binary image is used to get thetable borders
            thresh,img_bin = cv2.threshold(img, lower_thresh, upper_thresh, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
            imgname = self.check_for_dublicate("{}_cv_binary".format(fname), self.out_pth + "/images", ext = ".png" )
            cv2.imwrite(self.out_pth + '/images/' + imgname + '.png', img_bin)
            if self.debug:
                plotting = plt.imshow(img_bin,cmap='gray')
                plt.show()
            
            #inverting the image 
            img_bin = 255-img_bin
            imgname = self.check_for_dublicate("{}_cv_inverted".format(fname), self.out_pth + "/images", ext = ".png" )
            cv2.imwrite(self.out_pth + '/images/' + imgname + '.png', img_bin)
            
            if self.show_pics:
                #Plotting the image to see the output
                plotting = plt.imshow(img_bin,cmap='gray')
                plt.show()
            
            # countcol(width) of kernel as 100th of total width
            kernel_len = np.array(img).shape[1]//100
            # Defining a vertical kernel to detect all vertical lines of image 
            ver_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, kernel_len))
            # Defining a horizontal kernel to detect all horizontal lines of image
            hor_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_len, 1))
            # A kernel of 2x2
            kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
            
            #Use vertical kernel to detect and save the vertical lines in a jpg
            image_1 = cv2.erode(img_bin, ver_kernel, iterations=3)
            vertical_lines = cv2.dilate(image_1, ver_kernel, iterations=3)
            imgname = self.check_for_dublicate("{}_vertical".format(fname), self.out_pth + "/images", ext = ".jpg" )
            cv2.imwrite( self.out_pth + "/images/" + imgname +".jpg", vertical_lines)
            
            if self.show_pics:
                #Plot the generated image
                plotting = plt.imshow(image_1,cmap='gray')
                plt.show()
            
            #Use horizontal kernel to detect and save the horizontal lines in a jpg
            image_2 = cv2.erode(img_bin, hor_kernel, iterations=3)
            horizontal_lines = cv2.dilate(image_2, hor_kernel, iterations=3)
            imgname = self.check_for_dublicate("{}_horizontal".format(fname), self.out_pth + "/images", ext = ".jpg" )
            cv2.imwrite(self.out_pth + "/images/" + imgname + ".jpg", horizontal_lines)
            
            if self.show_pics:
                #Plot the generated image
                plotting = plt.imshow(image_2,cmap='gray')
                plt.show()
            
            # Combine horizontal and vertical lines in a new third image, with both having same weight.
            img_vh = cv2.addWeighted(vertical_lines, 0.5, horizontal_lines, 0.5, 0.0)
            #Eroding and thesholding the image
            img_vh = cv2.erode(~img_vh, kernel, iterations=2)
            thresh, img_vh = cv2.threshold(img_vh,128,255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
            imgname = self.check_for_dublicate("{}_img_vh".format(fname), self.out_pth + "/images", ext = ".jpg" )
            cv2.imwrite(self.out_pth + "/images/" + imgname + ".jpg", img_vh)
            # combining borders and text
            bitxor = cv2.bitwise_xor(img, img_vh)
            bitnot = cv2.bitwise_not(bitxor)
            self.bitnot_img = bitnot
            # self.show_pics = True
            if self.show_pics:
                #Plotting the generated image
                plotting = plt.imshow(bitxor,cmap='gray')
                plt.show()
                plotting = plt.imshow(bitnot,cmap='gray')
                plt.show()
            
            # Detect contours for following box detection
            contours, hierarchy = cv2.findContours(img_vh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
            
            def sort_contours(cnts, method="left-to-right"):
                # initialize the reverse flag and sort index
                reverse = False
                i = 0
                # handle if we need to sort in reverse
                if method == "right-to-left" or method == "bottom-to-top":
                    reverse = True
                # handle if we are sorting against the y-coordinate rather than
                # the x-coordinate of the bounding box
                if method == "top-to-bottom" or method == "bottom-to-top":
                    i = 1
                # construct the list of bounding boxes and sort them from top to
                # bottom
                boundingBoxes = [cv2.boundingRect(c) for c in cnts]
                (cnts, boundingBoxes) = zip(*sorted(zip(cnts, boundingBoxes),
                key=lambda b:b[1][i], reverse=reverse))
                # return the list of sorted contours and bounding boxes
                return (cnts, boundingBoxes)
            
            # Sort all the contours by top to bottom.
            contours, boundingBoxes = sort_contours(contours, method="top-to-bottom")
            
            # Creating a list of heights for all detected boxes - this may return false heights
            # heights = [boundingBoxes[i][3] for i in range(len(boundingBoxes))]
            
            # Get mean of heights
            # mean = np.mean(heights)
            
            # Create list box to store all boxes in  
            box = []
            heights = []
            
            # Get position (x,y), width and height for every contour and show the contour on image
            for c in contours:
                x, y, w, h = cv2.boundingRect(c)
                print("w", w, "h", h)
                
                # Limit the size of a cell
                if (w < 2000 and h < 700):
                    image = cv2.rectangle(img,(x,y),(x+w,y+h),(0,255,0), 4)
                    box.append([x,y,w,h])
                    heights.append(h)
            
            # Get mean of heights
            # This may cause false results if the table has different
            # cell-heights. Using actuall cell-height instead
            # mean_height = np.mean(heights)
            
            # if self.show_pics: 
            #     print("img with rectangles")
            #     plotting = plt.imshow(image,cmap='gray')
            #     plt.show()
            
            #Creating two lists to define row and column in which cell is located
            row=[]
            column=[]
            j = 0
            i = 0
            
            # Sorting the boxes to their respective row and column
            # if self.debug:
            #     print("len box", len(box) )  # box[1]
            #     for b in box:
            #         print(b)
            #     print("mean h", mean_height)
            #     input()

            
            for i in range(len(box)):    
                # box = [[x, y, w, h], ]
                # iterating over the number of rois found by cv2.cont
                # take care of change in rows and cols
                if(i==0):
                    column.append(box[i])
                    previous=box[i]    
                
                else:
                    # checking for changed row position
                    # opencv is going from 0 up
                    if(box[i][1] <= previous[1] + previous[3]/2):  # mean_height/2
                        column.append(box[i])
                        previous = box[i]            
                        # add to row if end box-list is reached
                        if(i == len(box) -1):
                            row.append(column)        
     
                    else:                        
                        row.append(column)
                        column=[]
                        previous = box[i]
                        column.append(box[i])
                        
            if self.debug:
                print("\nCOL:\n", len(column), column)
                print("\nROW:\n", len(row), row)
                # input()
            #calculating maximum number of cells
            countcol = 0
            for i in range(len(row)):
                countcol = len(row[i])
                if countcol > countcol:
                    countcol = countcol

            
            center = []
            if countcol != 0 and len(row) != 0:
                #Retrieving the center of each column
                center = [int(row[i][j][0]+row[i][j][2]/2) for j in range(len(row[i])) if row[0]]
                center=np.array(center)
                center.sort()
            if self.debug:
                print("\nCENTER:\n", center)
                
            #Regarding the distance to the columns center, the boxes are arranged in respective order
            finalboxes = []
            for i in range(len(row)):
                lis=[]
                for k in range(countcol):
                    lis.append([])
                for j in range(len(row[i])):
                    diff = abs(center-(row[i][j][0]+row[i][j][2]/4))
                    minimum = min(diff)
                    indexing = list(diff).index(minimum)
                    lis[indexing].append(row[i][j])
                finalboxes.append(lis)
                
            # leave the loop if table border was found
            if len(finalboxes) != 0:
                break
            else:
                # break
                if pcnt == 1:
                    break
                print("[INFO] Got no borders. Trying to sharpen the image")
                # # do some image transformation
                # try to sharpen the picture
                kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]])
                # img = cv2.filter2D(img, -1, kernel)
                img = cv2.filter2D(img, -1, kernel)
                img = cv2.GaussianBlur(img, (3,3),0)
                pcnt += 1
                               
        return finalboxes, countcol, row
            
        

    def get_ocrtext(self, finalboxes):
        # from every single cell/box the strings are extracted via pytesseract and stored in a list
        outer = []
        for i in range(len(finalboxes)):
            for j in range(len(finalboxes[i])):
                inner=''
                if(len(finalboxes[i][j])==0):
                    outer.append(' ')
                else:
                    for k in range(len(finalboxes[i][j])):
                        y,x,w,h = finalboxes[i][j][k][0],finalboxes[i][j][k][1], finalboxes[i][j][k][2],finalboxes[i][j][k][3]
                        finalimg = self.bitnot_img[x:x+h, y:y+w]
                        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 1))
                        border = cv2.copyMakeBorder(finalimg,2,2,2,2, cv2.BORDER_CONSTANT,value=[255,255])
                        resizing = cv2.resize(border, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
                        dilation = cv2.dilate(resizing, kernel,iterations=1)
                        erosion = cv2.erode(dilation, kernel,iterations=2)
                        # cv2.imshow("Erosion", erosion)
                        # cv2.waitKey(0)
                        out = pytesseract.image_to_string(erosion)
                        if(len(out)==0):
                            # If the rusult was zero, try different settings
                            out = pytesseract.image_to_string(erosion, config='--psm 3')
                        inner = inner +" "+ out
                        inner = self.text_postproc(inner)
                    
                    if self.debug:
                        print(inner)
                    outer.append(inner)
        return outer       
        
    
    def create_dataframe(self, ocr_data, countcol, row, out_name, rows_to_cols):      
        #Creating a dataframe of the generated OCR list
        arr = np.array(ocr_data)
        if not rows_to_cols:
            # Cols are cols, rows are rows
            dataframe = pd.DataFrame(arr.reshape(len(row), countcol))
        else:
            # Rows become cols, cols become rows
            dataframe = pd.DataFrame(arr.reshape(countcol, len(row)))
        print("\nDATAFRAME:\n", dataframe)
        data = dataframe.style.set_properties(align="left")
        # check for existing file
        out_name = self.check_for_dublicate(out_name, self.out_pth)
        #Converting it in a excel-file
        data.to_excel(self.out_pth + "\\{}.xlsx".format(out_name))
    
    def text_postproc(self, txt):
        txt = txt.replace("|","").replace("]","").replace("[","").replace("(","").replace(")","").replace("!","").replace('"','').replace("*","")
        return txt
    
    def check_for_dublicate(self, file_name, pth, ext = ".xlsx"):
        test_name = file_name
        if self.debug:
            print("[INFO] checking for dublicate names")
            print("--file:", test_name)
        fc = 1 
        cwd = os.getcwd()
        while True:                               
            if os.path.exists(cwd+"/" + pth +"/" + test_name + ext):
                test_name = file_name + " ({})".format(fc)
                fc += 1
            else:
                file_name = test_name
                if self.debug:
                    print("\tGot new filename:", file_name)
                return file_name
            
    
    def run(self, rows_to_cols = False):
        def process(file, out_name):
            # file is a string
            is_pth = True
            if self.convrt2table:
                # convert non table content to table
                file = self.mop.run(file)
                # file is not a string anymore, its an array
                is_pth = False
            print("[INFO] Getting boxes")
            fin_boxes, countcol, row = self.do_image_processing(file, out_name, is_pth)
            print("[INFO] getting ocr_text")
            ocr_text = self.get_ocrtext(fin_boxes)
            print("[INFO] Creating Excel file")
            self.create_dataframe(ocr_text, countcol, row, out_name, rows_to_cols)
        
        self.load_data()
        # for file in self.data:
        self.data.reverse()
        while len(self.data) != 0:
            thread_list = []
            for thr in range(self.cpu_threads):
                # check if data is still available
                if len(self.data) != 0:
                    file = self.data.pop()
                    print("File:", file)
                    out_name = file.split("/")[-1].split(".")[0]
                    # check if the file is a pdf
                    if re.search(".pdf", file):
                        self.convrt_pdf(file)
                    elif re.search(".gif", file) or re.search(".bmp", file):
                        pilimage = Image.open(file)
                        cv_image = np.array(pilimage)
                        self.cv_images.append(cv_image)
                    # process the pics got from pdf convert
                    if len(self.cv_images) != 0:
                        thr_cnt = 0
                        fn_add = len(self.cv_images)
                        while len(self.cv_images) != 0:
                            file = self.cv_images.pop()
                            # process(file, out_name)
                            # if thr_cnt != self.cpu_threads:
                            p = threading.Thread(target = process, args=(file, out_name + "_" + str(fn_add)))
                            # print("\nout_name addad:", out_name + "_" + str(fn_add))
                            
                            thread_list.append(p)
                            fn_add -= 1
                    else:
                        # process image file
                        # process(file, out_name)
                        p = threading.Thread(target = process, args=(file, out_name))
                        thread_list.append(p)
            
            # start the threads and wait for them
            thread_cnt = 0
            active_threads = []
            
            for p in thread_list:
                # if a pdf has more sites than self.cpu_threads, we will have to execute them by
                # bunches of self.cpu_threads
                # if thread_cnt != self.cpu_threads:
                # p = thread_list.pop()
                p.start()
                active_threads.append(p)
                if self.debug:
                    print("\n[INFO] Started Thread:", thread_cnt)
                    # input()
                thread_cnt += 1
                # else:
                if thread_cnt == self.cpu_threads:
                    # first: wait for the first threads to finsh ...
                    if self.debug:
                        print("Waiting for {} Threads to finisch".format(len(active_threads)))
                    for p in active_threads:
                        try:
                            p.join()
                        except Exception:
                            pass
                    # ... and reset
                    thread_cnt = 0
                    active_threads = []              
                           
            # wait for all left threads to finish
            if len(active_threads) != 0:
                if self.debug:
                    print("Waiting for {} left Threads to finisch".format(len(active_threads)))
            for p in active_threads:                
                try:
                    p.join()
                except Exception as e:
                    print(e.args)

class fileOCR_text_to_textfile:
    # Will save the ocr result as a textfile
    def __init__(self, tesseract_pth, input_path = "", output_path = "", use_all_files = True, cpu_threads = 4, 
                 debug = False, resize_first = False, invert_img = True):

        self.data_dir = input_path if not input_path == "" else "input"
        self.out_pth = output_path if not output_path == "" else "results"
        os.makedirs(self.out_pth, exist_ok=True)
        self.all_fs = use_all_files
        self.data = []
        self.debug = debug
        if self.debug:
            os.makedirs(self.out_pth + "/images", exist_ok=True)
        
        # if pdf file has more than one site, the converted images will be stored here
        self.cv_images = []
        
        # Multithreading
        self.cpu_threads = cpu_threads
        
        # save the pdf_name or gif name
        self.converted_file = ""
        
        # Save all ocr_results
        self.all_results = []
        
        # Resize before converting to binary image
        self.resize_first = resize_first
        self.resize_faktor = 2
        
        # inverting image: black text -> white text
        self.invert = invert_img
        
        # Check tesseract path
        self.tesseract_pwr = False
        if os.path.isfile(tesseract_pth):
            pytesseract.pytesseract.tesseract_cmd = tesseract_pth
            self.tesseract_pwr = True
        else:
            print("[ERROR] No valid tesseract.exe\nA B O R T I N G")
    
    def help(self):
        # Usage
        help_txt = """
        INITIALISIERUNG/CONSTRUCTOR
        ###########################
        
        tesseract_pth:\t 
        \tpath to tesseract.exe 
        
        input_path:\t 
        \tpath to the folder with the images to analyse
        \tNO SUBFOLDERS SUPPORTED ATM
        \tdefault: 'input' dir inside the folder of the file 
        
        output_path:\t 
        \tpath to the folder for the results  
        \tdefault: 'results' dir inside the folder of the file 
        \twill be created if no path is provided
        
        use_all_files:\t 
        \tif the input folder contains just images and pdfs
        \tyou can set this to True, otherwise it will filter the files
        \tdefault: True
        
        cpu_threads:\t 
        \thow many threads should be used
        \tdefault: 4
        
        debug:\t 
        \twill save the processed images to the results/images folder
        \tdefault: False
        
        resize_first:\t 
        \tif set to true, the input image will be first resized
        \tand then transformed into a binary image. this may lead to better
        \tbinary image quality and to better result. try at your own.
        \tdefault: False
        
        invert_img:\t 
        \tset to True if the fontcolor is dark and the background 
        \tis light. otherwise set to False.
        \tdefault: True
        
        RUN
        ##########################
        all_to_one_file:
        \tif set to True all results will be saved in one file. Otherwise
        \tevery input file will get its own resultfile
        \tdefault: False
        
        resize_faktor:
        \tinput images will be resized during preprocessing of ocr-detection.
        \tthis is the scale-factor 
        \tdefault: 2
        
        remove_empty_lines:
        \tif set to true all empty lines within the result will be removed.
        \tdefault: True
        """
        print(help_txt)
    
    def load_data(self):
        raw_data = os.listdir(self.data_dir)
        if not self.all_fs:
            # valid datatypes: jpg, jpeg, png, pdf
            vext = ["jpg", "jpeg", "png", "pdf"]
            
            for file in raw_data:
                ext = file.split(".")[-1]
                if ext in vext:
                    file = self.data_dir + "/" + file
                    self.data.append(file)
        else:
            for file in raw_data:
                file = self.data_dir + "/" + file
                self.data.append(file)
                
    def convrt_pdf(self, pdf_file):
        name = pdf_file.split("/")[-1].replace(".","_")
        self.converted_file = name
        pilimages = pdf2image.convert_from_path(pdf_file)
        cnt = 0
        for pilimage in pilimages:
            cv_image = np.array(pilimage) 
            # Convert RGB to BGR 
            cv_image = cv_image[:, :, ::-1]
            if self.debug:
                cv2.imwrite("temp/img_from_pdf_" + str(cnt)+".jpg", cv_image)
                cnt += 1
            self.cv_images.append(cv_image)
            
    def do_image_processing(self, file, fname, file_is_path = True, lower_thresh = 64, upper_thresh = 255):      
        
        if file_is_path:
            if type(file) == str:            
                img = cv2.imread(file, 0)
                # print(type(img))
            else:
                # got a converted pdf image
                # print(type(file))
                os.makedirs("temp", exist_ok=True)
                temp_name = "temp_{}".format(fname)
                fname = self.check_for_dublicate(temp_name, "temp", ext = ".png")
                cv2.imwrite("temp/{}.png".format(fname), file)
                img = cv2.imread("temp/{}.png".format(fname), 0)
                # try:
                #     os.remove("temp/{}.png".format(fname))
                # except Exception:
                #     pass
        else:
            # Image aus ocr_preps erhalten
            img = file
          
        # thresholding the image to a binary image
        # binary image is used to get thetable borders
        if self.resize_first:
            img = cv2.resize(img, None, fx= self.resize_faktor, fy= self.resize_faktor, interpolation=cv2.INTER_CUBIC)
        thresh,img_bin = cv2.threshold(img, lower_thresh, upper_thresh, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        if self.debug:
            imgname = self.check_for_dublicate("{}_cv_binary".format(fname), self.out_pth + "/images", ext = ".png" )
            cv2.imwrite(self.out_pth + '/images/' + imgname + '.png', img_bin)
        
        #inverting the image
        if self.invert:
            img_bin = 255-img_bin
        if self.debug:
            imgname = self.check_for_dublicate("{}_cv_inverted".format(fname), self.out_pth + "/images", ext = ".png" )
            cv2.imwrite(self.out_pth + '/images/' + imgname + '.png', img_bin)
        
        return img_bin
    
    def get_ocrtext(self, img_bin, dil_iter = 3, er_iter = 2, resize_faktor = 2, kernel_size = (2, 1)):
        # Get the ocr result
        # outer=[]
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, kernel_size)
        if not self.resize_first:
            resizing = cv2.resize(img_bin, None, fx= resize_faktor, fy= resize_faktor, interpolation=cv2.INTER_CUBIC)
        else:
            resizing = img_bin
        dilation = cv2.dilate(resizing, kernel,iterations = dil_iter)
        erosion = cv2.erode(dilation, kernel,iterations = er_iter)
        # if self.debug:
            # cv2.imshow("Erosion", erosion)
            # cv2.waitKey(0)
        out = pytesseract.image_to_string(erosion)
        if(len(out) == 0):
            # If the rusult was zero, try different settings
            out = pytesseract.image_to_string(erosion, config='--psm 3')

        return out
    
    def text_postproc(self, txt):
        txt = txt.replace("|","").replace("]","").replace("[","").replace("(","").replace(")","").replace("!","").replace('"','').replace("*","")
        return txt
    
    def check_for_dublicate(self, file_name, pth, ext = ".txt"):
        test_name = file_name
        if self.debug:
            print("[INFO] checking for dublicate names")
            print("--file:", test_name)
        fc = 1 
        cwd = os.getcwd()
        while True:                               
            if os.path.exists(cwd+"/" + pth +"/" + test_name + ext):
                test_name = file_name + " ({})".format(fc)
                fc += 1
            else:
                file_name = test_name
                if self.debug:
                    print("\tGot new filename:", file_name)
                return file_name
            
    def save_to_file(self, ocr_data, out_name, remove_empty_lines):
        # save the ocr_result to a file
        # if all_to_one_file is True, all result will be stored in one file
        # if not,  every result get it's own file
        if len(self.all_results) != 0:
            # Get a filename
            fname = self.check_for_dublicate( out_name, self.out_pth)
            print(fname)
            print(self.out_pth + "/" + fname + ".txt")
            # create the file
            fn = self.out_pth + "/" + fname + ".txt"
            with open(fn, "w+") as f:
            # Write to file
                for res in self.all_results:
                    # writing results and remove empty lines if wished
                    f.write(res if not remove_empty_lines else res.replace("\n\n","\n") + "\n")
                    # empty line between results
                    f.write("\n")
            
        else:
            fname = self.check_for_dublicate( out_name, self.out_pth) 
            with open(self.out_pth + "/" + fname + ".txt", "w+") as f:
                f.write(ocr_data + "\n")
               
    def run(self, all_to_one_file = False, resize_faktor = 2, remove_empty_lines = True):
        if self.tesseract_pwr:
            # update the default resize_faktor
            self.resize_faktor = resize_faktor
            
            def process(file, out_name):
                # file is a string
                is_pth = True
                print("[INFO] Processing Image")
                img_bin = self.do_image_processing(file, out_name, is_pth)
                print("[INFO] getting ocr_text")
                ocr_text = self.get_ocrtext(img_bin, resize_faktor = self.resize_faktor)       
                if not all_to_one_file:
                    print("[INFO] Saving file")
                    self.save_to_file(ocr_text, out_name, remove_empty_lines)
                else:
                    # Add the filename to the result to save belonging
                    self.all_results.append(out_name + ":\n" + ocr_text)
            
            self.load_data()
            # for file in self.data:
            self.data.reverse()
            while len(self.data) != 0:
                thread_list = []
                for thr in range(self.cpu_threads):
                    # check if data is still available
                    if len(self.data) != 0:
                        file = self.data.pop()
                        print("File:", file)
                        out_name = file.split("/")[-1].split(".")[0]
                        # check if the file is a pdf
                        if re.search(".pdf", file):
                            self.convrt_pdf(file)
                        elif re.search(".gif", file) or re.search(".bmp", file):
                            pilimage = Image.open(file)
                            cv_image = np.array(pilimage)
                            self.cv_images.append(cv_image)
                        # process the pics got from pdf convert
                        if len(self.cv_images) != 0:
                            thr_cnt = 0
                            fn_add = len(self.cv_images)
                            while len(self.cv_images) != 0:
                                file = self.cv_images.pop()
                                # process(file, out_name)
                                # if thr_cnt != self.cpu_threads:
                                p = threading.Thread(target = process, args=(file, out_name + "_" + str(fn_add)))
                                # print("\nout_name addad:", out_name + "_" + str(fn_add))
                                
                                thread_list.append(p)
                                fn_add -= 1
                        else:
                            # process image file
                            # process(file, out_name)
                            p = threading.Thread(target = process, args=(file, out_name))
                            thread_list.append(p)
                
                # start the threads and wait for them
                thread_cnt = 0
                active_threads = []
                
                for p in thread_list:
                    # if a pdf has more sites than self.cpu_threads, we will have to execute them by
                    # bunches of self.cpu_threads
                    # if thread_cnt != self.cpu_threads:
                    # p = thread_list.pop()
                    p.start()
                    active_threads.append(p)
                    if self.debug:
                        print("\n[INFO] Started Thread:", thread_cnt)
                        # input()
                    thread_cnt += 1
                    # else:
                    if thread_cnt == self.cpu_threads:
                        # first: wait for the first threads to finsh ...
                        if self.debug:
                            print("Waiting for {} Threads to finisch".format(len(active_threads)))
                        for p in active_threads:
                            try:
                                p.join()
                            except Exception:
                                pass
                        # ... and reset
                        thread_cnt = 0
                        active_threads = []              
                               
                # wait for all left threads to finish
                if len(active_threads) != 0:
                    if self.debug:
                        print("Waiting for {} left Threads to finisch".format(len(active_threads)))
                for p in active_threads:                
                    try:
                        p.join()
                    except Exception as e:
                        print(e.args)
            
            # Save results
            if all_to_one_file:
                self.save_to_file("", out_name, remove_empty_lines)

# =============================================================================
# Copy Utils
# =============================================================================
def remove_linebreaks(delimeter = None):
    # just a function to get the clipbpard content, remove the
    # linebreaks und return the result to clipboard
    if delimeter is None:
        sep = ""
    else:
        sep = delimeter
    txt = pyperclip.paste()
    
    txt = txt.replace("\r\n", sep + " " ).replace("\n", sep + " " ).replace("  ", " ")
    pyperclip.copy(txt)
    
def linebreak_to_excelcols(seperator = "\t"):
    txt = pyperclip.paste()
    if re.search("\r\n", txt) is not None:
        txt = txt.replace("\r\n", seperator)
    else:
        txt = txt.replace("\n", seperator)
    txt = txt.replace(seperator+seperator, seperator).replace("  ", " ")
    pyperclip.copy(txt)



if __name__ == "__main__":
    # scrn = screenshot()
    # scrn.run()
    
    # prep = Prepare(search_vertical = False)
    # # image = cv2.imread("....jpg")
    # # image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    # # cv2.imshow("image", image)
    # # input()
    # prep.run("....jpg") 
    # cv2.destroyAllWindows()
    
    # ocr2c = screenshotOCR_to_clipboard(r'...\tesseract.exe')
    # result_as_line = ocr2c.ocr_to_oneline()
    # print("Result oneline:", result_as_line)
    
    # result_as_rows = ocr2c.ocr_to_rows()
    # print("Result rows:", result_as_rows)
    
    # result_as_cols = ocr2c.ocr_to_cols()
    # print("Result cols:", result_as_cols)
    # ocr2c.copy_to_clipboard()
    
    # result_as_df = ocr2c.ocr_table_to_table()
    # print("result_as_df:", result_as_df)
    
    # ocr2c.copy_to_clipboard()
    # cv2.destroyAllWindows()
    
    # start =  time.time()
    # # Set your options here
    # input_dir = "input"
    # output_dir = "ocr_results"
    # number_of_threads = multiprocessing.cpu_count()
    # tesseract_pth = r'...\tesseract.exe'
    # ocr = fileOCR_table_to_xlsx(tesseract_pth, input_dir, output_dir, cpu_threads = number_of_threads, debug = True)
    # ocr.run()
    # print("DONE!\nThis took: {} seconds".format(time.time()-start))
    
    # start =  time.time()
    # # Set your options here
    # input_dir = "input"
    # output_dir = "ocr_results"
    # number_of_threads = multiprocessing.cpu_count()
    # tesseract_pth = r'...\tesseract.exe'
    # ocr = fileOCR_text_to_textfile(tesseract_pth, input_dir, output_dir, cpu_threads = number_of_threads, debug = True)
    # ocr.help()
    # ocr.run(all_to_one_file = True, resize_faktor = 2)
    # print("DONE!\nThis took: {} seconds".format(time.time()-start))
    

# Copyright 2021 Agraj Abhishek. All Rights Reserved.
#
# Licensed under the MIT License, (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     https://github.com/aabhi1/Merge-Join-PDF/blob/main/LICENSE
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
# ==============================================================================

import os, cv2, sys
import tkinter as tk
from tkinter import Menu, ttk, TOP, filedialog, messagebox
from PIL import ImageTk, Image
from fpdf import FPDF
from datetime import datetime
import tkinter.font as font
from tkinter.ttk import *
from tkinter import StringVar
import fitz
class MainApp(tk.Tk):
    def sel(self):
        self.swapV.append(self.v.get())
        if len(self.swapV) == 2:
            xc = self.afile[int(self.swapV[0])]
            self.afile[int(self.swapV[0])] = self.afile[int(self.swapV[1])]
            self.afile[int(self.swapV[1])] = xc
            self.swapV = []
            i = j = 0
            for imgP in range(0, len(self.afile)):
                self.img = ImageTk.PhotoImage(Image.open(self.afile[imgP]).resize((170, 170)))
                photo = tk.Label(self, image=self.img)
                photo.image = self.img
                photo.grid(row=i, column=j)
                j += 1
                if j % 7 == 0:
                    j = 0
                    i += 2
            messagebox.showinfo('Success','Pages Swapped Successfully !')

    def fimageUpload(self):
        self.v = StringVar(self, "1")

        text_file_extensions = ['*.pdf', '*.pdF', '*.pDF', '*.Pdf', '*.PDf', '*.PDF', '*.pDf', '*.jpg', '*.jpeg',
                                '*.png']
        image_file_extensions = ['*.jpg', '*.JPG', '*.png', '*.PNG', '*.gif', '*.GIF', '*.jpeg', '*.JPEG', '*.tif',
                                '*.TIF','*.bmp', '*.BMP']
        ftypes = [('Pdf Files', text_file_extensions), ('Image Files', image_file_extensions)]
        self.filename = list(filedialog.askopenfilenames(initialdir="/", title="Select file",
                                                   filetypes=ftypes))
        for selF in range(0,len(self.filename)):
            i = j = 0
            if self.filename[selF].lower().endswith(".jpg") or self.filename[selF].lower().endswith(".PNG") or self.filename[selF].lower().endswith(
                ".png") or self.filename[selF].lower().endswith(".jpeg") or self.filename[selF].lower().endswith(".gif") or \
                self.filename[selF].lower().endswith(".bmp") or self.filename[selF].lower().endswith(".tif"):

                self.afile.append(self.filename[selF])
                for imgP in range(0, len(self.afile)):
                    try:
                        self.img = ImageTk.PhotoImage(Image.open(self.afile[imgP]).resize((170, 170)))
                        photo = tk.Label(self, image=self.img)
                        photo.image = self.img
                        photo.grid(row=i, column=j)
                        Radiobutton(self, text="Frame: " + str(imgP), variable=self.v, value=imgP, width="25",
                                    command=self.sel).grid(row=i + 1, column=j)
                        j += 1
                        if j % 7 == 0:
                            j = 0
                            i += 2
                    except Exception as e:
                        messagebox.showerror("Error","Image file type error "+str(e))
                        del self.afile[imgP]
                        pass

            elif self.filename[selF].lower().endswith(".pdf"):
                try:
                    file = self.filename[selF]
                    doc = fitz.open(file)
                    for page_index in range(len(doc)):
                        first_page = doc[page_index]
                        image_matrix = fitz.Matrix(fitz.Identity)
                        image_matrix.preScale(3, 3)
                        pix = first_page.getPixmap(alpha=False, matrix=image_matrix)
                        output_img = self.tempdirpath+ "/" + file.split("/")[-1].split(".pdf")[0] + str(page_index) + ".png"
                        pix.writePNG(output_img)
                        self.afile.append(output_img)

                    for imgP in range(0, len(self.afile)):
                        self.img = ImageTk.PhotoImage(Image.open(self.afile[imgP]).resize((170, 170)))
                        photo = tk.Label(self, image=self.img)
                        photo.image = self.img
                        photo.grid(row=i, column=j)
                        Radiobutton(self, text="Frame: " + str(imgP), variable=self.v, value=imgP, width="25",
                                command=self.sel).grid(row=i + 1, column=j)
                        j += 1
                        if j % 7 == 0:
                            j = 0
                            i += 2
                except Exception as e:
                    messagebox.showerror("Error","Image file type error "+str(e))
                    del self.filename[selF]
                    pass

    def createPdf(self, pdfStyle):
        im = []
        for i in range(0, len(self.afile)):
            if pdfStyle == 1:
                img = Image.open(self.afile[i]).resize((2480, 3508))
            elif pdfStyle == 0:
                img = Image.open(self.afile[i]).resize((3508, 2480))
            if img.mode == 'RGBA':
                img = img.convert('RGB')
            im.append(img)  # resize((827,1170 )))
        if pdfStyle == 1:
            im1 = Image.open(self.afile[0]).resize((2480, 3508))
        elif pdfStyle == 0:
            im1 = Image.open(self.afile[0]).resize((3508, 2480))
        if im1.mode == 'RGBA':
            im1 = im1.convert('RGB')

        dtnow = datetime.now().strftime("%Y%m%d-%H%M%S")
        pdf1_filename = dtnow + ".pdf"
        im[0].save(pdf1_filename, "PDF", quality = 10, resolution = 90, optimize = True, save_all = True, append_images = im[1:])
        messagebox.showinfo('Success', 'PDF Generated with filename '+str(pdf1_filename)+ ' !')

    def pdfUpload(self):
        text_file_extensions = ['*.pdf', '*.pdF', '*.pDF', '*.Pdf', '*.PDf', '*.PDF', '*.pDff']
        ftypes = [('Pdf Files', text_file_extensions)]
        self.pdffilename = list(filedialog.askopenfilenames(initialdir="/", title="Select file",
                                                   filetypes=ftypes))
        for selF in range(0, len(self.pdffilename)):
            self.pdfafile.append(self.pdffilename[selF])
        messagebox.showinfo("Success", "Uploaded " + str(len(self.pdffilename)) + " PDF file !")

    def pdfJoin(self):
        try:
            messagebox.showinfo("Info", "Merging " + str(len(self.pdffilename)) + " file !")
            if len(self.pdfafile) > 0:
                fname = self.pdfafile[0]
                docname = fname
                doc11 = fitz.open(fname)
                toc11 = doc11.getToC(False)
                pc1 = len(doc11)

                for i in range (1,len(self.pdfafile)):
                    fname = self.pdfafile[i]
                    doc2 = fitz.open(fname)
                    toc2 = doc2.getToC(False)

                    new_toc2 = []
                    for line in toc2:
                        line[2] += pc1
                        new_toc2.append(line)

                    doc11.insertPDF(doc2)  # append file 2 to file 1
                    doc11.setToC(toc11 + new_toc2)  # set table of contents for the result
                    i +=2
                dtnow = datetime.now().strftime("%Y%m%d-%H%M%S")
                pdf1_filename = dtnow + ".pdf"
                doc11.save(pdf1_filename)  # save result to new file
                messagebox.showinfo("Success", "Saved as " + pdf1_filename + " filename !")
            self.pdffilename = []
            self.pdfafile = []
        except Exception as e:
            messagebox.showerror("Error", "Image file type error " + str(e))
            self.pdffilename = []
            self.pdfafile = []


    def __init__(self):
        tk.Tk.__init__(self)
        self.configure(background="black")
        self.x = self.y = 0

        if getattr(sys, 'frozen', False):
            application_path = sys._MEIPASS
        elif __file__:
            application_path = os.path.dirname(__file__)

        iconFile = 'logo.ico'

        logoFile = 'logo.png'

        self.iconbitmap(default=os.path.join(application_path, iconFile))
        tmpdir = "temporary"
        self.tempdirpath = os.path.join(application_path, tmpdir)
        if not os.path.exists(self.tempdirpath):
            os.mkdir(self.tempdirpath)

        self.swapV = []
        self.afile = []
        self.pdfafile = []
        self.title("JPG2PDF - Beta")

        myFont = font.Font(weight="bold")
        menubar = Menu(self)
        self.filemenu = Menu(menubar, tearoff=0)
        self.filemenu.add_command(label="Browse Image File", command=lambda: self.fimageUpload(), accelerator="Ctrl+O",
                                  activebackground='#e4642c')
        self.filemenu.add_separator()

        self.filemenu.add_command(label="Create PDF - Landscape", command=lambda: self.createPdf(0), accelerator="Ctrl+S",
                                  activebackground='#e4642c')

        self.filemenu.add_command(label="Create PDF - Potrait", command=lambda: self.createPdf(1), accelerator="Ctrl+S",
                                  activebackground='#e4642c')

        self.filemenu.add_separator()

        self.filemenu.add_command(label="Exit", command=self.quit, accelerator="Ctrl+Z", activebackground='#e4642c')
        menubar.add_cascade(label="IMG2PDG", menu=self.filemenu)
        self.config(menu=menubar)

        self.helpmenu = Menu(menubar, tearoff=0)

        self.helpmenu.add_command(label="Upload PDF", command=lambda: self.pdfUpload(), accelerator="Ctrl+I",
                                  activebackground='#e4642c')

        self.helpmenu.add_command(label="Join PDF", command=lambda: self.pdfJoin(), accelerator="Ctrl+I",
                                  activebackground='#e4642c')
        menubar.add_cascade(label="PDF2PDF", menu=self.helpmenu)
        self.config(menu=menubar)

        self.helpmenu = Menu(menubar, tearoff=0)
        self.helpmenu.add_command(label="Help", command=lambda: self.fimageUpload(self.afile), accelerator="Ctrl+H",
                                  activebackground='#e4642c')
        menubar.add_cascade(label="Help", menu=self.helpmenu)
        self.config(menu=menubar)
        img = ImageTk.PhotoImage(Image.open(os.path.join(application_path, logoFile)).resize((170, 170)))

#        img = ImageTk.PhotoImage(Image.open(self.tempdirpath+'\logo.PNG').resize((170, 170)))
        logo = tk.Label(self, image=img)
        logo.image = img
        logo.grid(row=0, column=0)


if __name__ == "__main__":
    app = MainApp()
    app.mainloop()

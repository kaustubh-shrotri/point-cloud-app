import tkinter as tk
import tkinter.ttk as ttk
import open3d as o3d
import numpy as np
from tkinter import filedialog,IntVar,messagebox
import os
import simplejson as json
import xlsxwriter
import time



class NewprojectApp:
    def __init__(self, master=None):
        # build ui
        self.PointCloudApp = ttk.Frame(master)
        self.label_1 = ttk.Label(self.PointCloudApp)
        self.label_1.configure(anchor='w', font='{Helvetica} 14 {}', text='Welcome to the Point Cloud Application.')
        self.label_1.grid(columnspan='5')
        self.entry_1 = ttk.Entry(self.PointCloudApp)
        self.entry_1.grid(ipadx='30', row='2')
        self.entry_11 = ttk.Entry(self.PointCloudApp)
        self.button_1 = ttk.Button(self.PointCloudApp, command=lambda:self.fileopen(self.entry_1,self.entry_11))
        self.button_1.configure(text='Select File')
        self.button_1.grid(column='1', padx='2', row='2')
        self.button_2 = ttk.Button(self.PointCloudApp, command=lambda:self.clear_text(self.entry_1,self.entry_11))
        self.button_2.configure(text='Remove')
        self.button_2.grid(column='2', padx='2', row='2')
        self.label_2 = ttk.Label(self.PointCloudApp)
        self.label_2.configure(font='{Times} 11 {}', text='Please Select the Point Cloud file.')
        self.label_2.grid(row='1', sticky='w')
        self.button_3 = ttk.Button(self.PointCloudApp, command =self.view_point_cloud)
        self.button_3.configure(state='disabled', text='View Point Cloud')
        self.button_3.grid(column='1', columnspan='2', ipadx='28', pady='5', row='3')
        self.separator_1 = ttk.Separator(self.PointCloudApp)
        self.separator_1.configure(orient='horizontal')
        self.separator_1.grid(column='0', columnspan='3', ipadx='180', pady='5', row='4')
        self.label_3 = ttk.Label(self.PointCloudApp)
        self.label_3.configure(font='{Times} 11 {}', text='Compute distance between two points')
        self.label_3.grid(column='0', row='5', sticky='w')
        self.entry_2 = ttk.Entry(self.PointCloudApp)
        self.var2 = tk.IntVar()
        self.entry_2.configure(textvariable=self.var2)
        self.entry_2.grid(column='0', ipadx='30', row='6')
        self.entry_2.rowconfigure('6', minsize='0')
        self.label_4 = ttk.Label(self.PointCloudApp)
        self.label_4.configure(font='{Times} 10 {}', text='and')
        self.label_4.grid(column='0', row='7')
        self.entry_3 = ttk.Entry(self.PointCloudApp)
        self.var3 = tk.IntVar()
        self.entry_3.configure(textvariable=self.var3)
        self.entry_3.grid(column='0', ipadx='30', row='8')
        self.button_4 = ttk.Button(self.PointCloudApp)
        self.button_4.configure(text='Measure', command=self.measure_distance)
        self.button_4.grid(column='1', columnspan='2', ipadx='10', ipady='10', row='6', rowspan='2')
        self.text_1 = tk.Text(self.PointCloudApp)
        self.var4 = tk.DoubleVar()
        self.text_1.configure( height='0', width='10')
        self.text_1.grid(column='0', pady='5', row='9')
        self.label_5 = ttk.Label(self.PointCloudApp)
        self.label_5.configure(text='Distance: ')
        self.label_5.grid(column='0', padx='5', pady='5', row='9', sticky='w')
        self.label_6 = ttk.Label(self.PointCloudApp)
        self.label_6.configure(text='mm')
        self.label_6.grid(column='0', padx='30', pady='5', row='9', sticky='e')
        self.separator_2 = ttk.Separator(self.PointCloudApp)
        self.separator_2.configure(orient='horizontal')
        self.separator_2.grid(column='0', columnspan='3', ipadx='180', pady='5', row='10')
        self.label_7 = ttk.Label(self.PointCloudApp)
        self.label_7.configure(font='{Times} 11 {}', text='Calculate Coordinates of the holes.')
        self.label_7.grid(column='0', row='11', sticky='w')
        self.text_2 = tk.Text(self.PointCloudApp)
        self.text_2.configure(height='5', width='30')
        self.text_2.grid(column='0', columnspan='2', ipadx='10', row='12')
        self.button_5 = ttk.Button(self.PointCloudApp, command=self.find_centroid)
        self.button_5.configure(text='Calculate')
        self.button_5.grid(column='2', ipady='10', row='12')
        self.separator_3 = ttk.Separator(self.PointCloudApp)
        self.separator_3.configure(orient='horizontal')
        self.separator_3.grid(column='0', columnspan='3', ipadx='180', pady='5', row='13')
        self.label_8 = ttk.Label(self.PointCloudApp)
        self.label_8.configure(font='{Times} 11 {}', text='Export Data')
        self.label_8.grid(column='0', row='14', sticky='w')
        self.entry_5 = ttk.Entry(self.PointCloudApp)
        self.entry_5.grid(column='0', ipadx='30', row='15')
        self.button_6 = ttk.Button(self.PointCloudApp, command=lambda:self.folderopen(self.entry_5))
        self.button_6.configure(text='Select Folder')
        self.button_6.grid(column='1', padx='2', row='15')
        self.button_7 = ttk.Button(self.PointCloudApp, command= self.export_data)
        self.button_7.configure(text='Export')
        self.button_7.grid(column='2', padx='2', row='15')
        self.progressbar_1 = ttk.Progressbar(self.PointCloudApp)
        self.progressbar_1.configure(length='370', orient='horizontal')
        self.progressbar_1.grid(column='0', columnspan='3', pady='5', row='16')
        self.separator_4 = ttk.Separator(self.PointCloudApp)
        self.separator_4.configure(orient='vertical')
        self.separator_4.grid(column='3', ipady='200', padx='10', row='0', rowspan='17')
        self.text_3 = tk.Text(self.PointCloudApp)
        self.text_3.configure(font='{calibri} 10 {}', height='25', insertwidth='50', width='50', state='disabled')
        _text_ = '''Help Guide!\n
--Load a point cloud file by clicking on \'Select File\'. \n
--You can view the loaded file by clicking on \'View Point Cloud\'. \n 
--To measure the distance between two points click on\n \'Measure\'. \n
--To calculate the holes coordinates click on \'Calculate\'. \n
--After doing the calculations and/or measurements, click on\n \'Export\' to export generated data. \n
--You can also select a specific folder to store exported data.\n Click on \'Select Folder\'. \n
--The exported file will be in .xlsx format with the time stamp of\n file generation. '''
        self.text_3.configure(state='normal')
        self.text_3.insert('0.0', _text_)
        self.text_3.configure(state='disabled')
        self.text_3.grid(column='4', row='1', rowspan='17')
        self.scrollbar_2 = ttk.Scrollbar(self.PointCloudApp, command=self.text_3.yview)
        self.scrollbar_2.configure(orient='vertical')
        self.scrollbar_2.grid(column='5', ipady='164', row='1', rowspan='17', sticky='e')
        self.text_3['yscrollcommand'] = self.scrollbar_2.set
        self.message_2 = tk.Message(self.PointCloudApp)
        self.var = tk.StringVar('')
        self.message_2.configure(font='{calibri} 8 {}',textvariable=self.var, width='200')
        self.message_2.grid(column='0', row='3')
        self.PointCloudApp.configure(cursor='arrow', height='500', takefocus=False, width='400')
        self.PointCloudApp.pack(side='top')

        # Main widget
        self.mainwindow = self.PointCloudApp
        self.measure_dist_lst = []
        self.centroid_lst = []
        self.out_dir = None


    def run(self):
        self.mainwindow.mainloop()
    
    def view_point_cloud(self):
        text = '''-- Mouse view control --
     Left button + drag         : Rotate.
     Ctrl + left button + drag  : Translate.
     Wheel button + drag        : Translate.
     Shift + left button + drag : Roll.
     Wheel                      : Zoom in/out.

-- Keyboard view control --
     [/]          : Increase/decrease field of view.
     R            : Reset view point.
     Ctrl/Cmd + C : Copy current view status into the clipboard.
     Ctrl/Cmd + V : Paste view status from clipboard.

-- General control --
     Q, Esc       : Exit window.
     H            : Print help message.
     P, PrtScn    : Take a screen capture.
     D            : Take a depth capture.
     O            : Take a capture of current rendering settings.
     Alt + Enter  : Toggle between full screen and windowed mode.

-- Render mode control --
     L            : Turn on/off lighting.
     +/-          : Increase/decrease point size.
     Ctrl + +/-   : Increase/decrease width of geometry::LineSet.
     N            : Turn on/off point cloud normal rendering.
     S            : Toggle between mesh flat shading and smooth shading.
     W            : Turn on/off mesh wireframe.
     B            : Turn on/off back face rendering.
     I            : Turn on/off image zoom in interpolation.
     T            : Toggle among image render:
                    no stretch / keep ratio / freely stretch.'''
        self.text_3.configure(state='normal')
        self.text_3.delete(3.0,tk.END)
        self.text_3.insert('3.0', text)
        self.text_3.configure(state='disabled')
        self.PointCloudApp.update_idletasks()
        o3d.visualization.draw_geometries([self.pcd], width=1080, height=700)
        
    def fileopen(self, x,y):
        x.delete(0, 'end')
        y.delete(0, 'end')
        self.excel1=filedialog.askopenfilename()
        y.insert(tk.END,self.excel1)
        self.excel2 = os.path.basename(self.excel1)
        x.insert(tk.END,self.excel2)
        self.pcd = o3d.io.read_point_cloud(self.excel1)
        self.var.set(f'{self.excel2} loaded successfully.' )
        self.button_3['state']=tk.NORMAL
        self.iteration = 1
        self.pcd_asarray = np.asarray(self.pcd.points)
        self.measure_dist_lst = []
        self.centroid_lst = []
    
    def folderopen(self,x):
        x.delete(0, 'end')
        self.out_dir = filedialog.askdirectory()
        x.insert(tk.END,self.out_dir)
        
    def clear_text(self,x,y):
        x.delete(0, 'end')
        y.delete(0, 'end')
        self.var.set('Point Cloud file removed.')
        self.button_3['state']=tk.DISABLED
        self.var2.set('')
        self.var3.set('')
        self.text_1.delete(1.0,tk.END)
        self.text_2.delete(1.0,tk.END)
        self.entry_5.delete(0,'end')
        
        
    def measure_distance(self):
        print(type(self.entry_2.get()),self.entry_2.get())
        if self.entry_2.get() == '0' or self.entry_3.get() == '0':
            self.eucd_distance()
        else:
            self.distance_bw_holes()
            
    def distance_bw_holes(self):
        self.point1 = np.asarray(json.loads(self.entry_2.get()))
        self.point2 = np.asarray(json.loads(self.entry_3.get()))
        self.distance = np.linalg.norm(self.point1-self.point2)
        self.text_1.delete(1.0,2.0)
        self.text_1.insert(1.0,self.distance)
        self.measure_dist_lst.append([self.point1, self.point2, self.distance])
        
    def eucd_distance(self):
        text = '''\n1) Please pick two correspondences using [shift + left click]
           Press [shift + right click] to undo point picking
2) After picking points, press 'Q' to close the window'''
        self.text_3.configure(state='normal')
        self.text_3.delete(2.0,tk.END)
        self.text_3.insert('2.0', text)
        self.text_3.configure(state='disabled')
        self.PointCloudApp.update_idletasks()
        self.vis = o3d.visualization.VisualizerWithEditing()
        self.vis.create_window(width=1080, height=700)
        self.vis.add_geometry(self.pcd)
        self.vis.run()  # user picks points
        self.vis.destroy_window()
        # print("")
        self.points = self.vis.get_picked_points()
        self.point1 = self.pcd_asarray[self.points[0]]
        self.var2.set(self.point1)
        self.point2 = self.pcd_asarray[self.points[1]]
        self.var3.set(self.point2)
        self.distance = np.linalg.norm(self.point1-self.point2)
        self.text_1.delete(1.0,2.0)
        self.text_1.insert(1.0,self.distance)
        self.measure_dist_lst.append([self.point1, self.point2, self.distance])
        
        # return print(f'Distance between point {a} and point {b} is {distance}')
        
    def find_centroid(self):
        text = '''\n1) Please pick at least three correspondences using [shift + left click]
           Press [shift + right click] to undo point picking
2) After picking points, press 'Q' to close the window'''
        self.text_3.configure(state='normal')
        self.text_3.delete(2.0,tk.END)
        self.text_3.insert('2.0', text)
        self.text_3.configure(state='disabled')
        self.PointCloudApp.update_idletasks()
        self.vis = o3d.visualization.VisualizerWithEditing()
        self.vis.create_window(width=1080, height=700)
        self.vis.add_geometry(self.pcd)
        self.vis.run()  # user picks points
        self.vis.destroy_window()
        # print("")
        self.points = self.pcd_asarray[self.vis.get_picked_points()]
        self.centroid = np.mean(self.points, axis=0)
        self.holes_cood = f'Hole {self.iteration} co-ordinaters are: {self.centroid}'
        self.text_2.insert(tk.END, self.holes_cood+'\n')
        self.centroid_lst.append(['Hole '+str(self.iteration), self.centroid])
        self.iteration +=1
        
        # return print(f'The hole coordinate is {self.centroid}')
        
    def export_data(self):
        time_str=time.strftime("%y_%m_%d__%H_%M_%S")
        if self.out_dir:
            file = self.out_dir+'/report_'+time_str+'.xlsx'
            workbook = xlsxwriter.Workbook(file)
        else:
            file = os.path.dirname(self.excel1)+'/report_'+time_str+'.xlsx'
            workbook = xlsxwriter.Workbook(file)
        
        # workbook = xlsxwriter.Workbook('report_'+time_str+'.xlsx')
        if self.measure_dist_lst:
            self.progressbar_1['value']=30
            self.PointCloudApp.update_idletasks()
            worksheet1 = workbook.add_worksheet(name='distance measurement')
            worksheet1.write(0,0,'Sr. No.')
            worksheet1.write(0,1,'Point1')
            worksheet1.write(0,2,'Point2')
            worksheet1.write(0,3,'Distance Point1 to Point2')
            row = 1
            for i, j in enumerate(self.measure_dist_lst):
                worksheet1.write(row,0,i+1)
                worksheet1.write(row,1,str(j[0]))
                worksheet1.write(row,2,str(j[1]))
                worksheet1.write(row,3,str(j[2]))
                row +=1
        if self.centroid_lst:
            self.progressbar_1['value']=60
            self.PointCloudApp.update_idletasks()
            worksheet2 = workbook.add_worksheet(name='Holes Centroid')
            worksheet2.write(0,0,'Sr. No.')
            worksheet2.write(0,1,'Holes')
            worksheet2.write(0,2,'Co-ordinates')
            row=1
            for i, j in enumerate(self.centroid_lst):
                worksheet2.write(row,0,i+1)
                worksheet2.write(row,1,str(j[0]))
                worksheet2.write(row,2,str(j[1]))
                row+=1
            
        if not (self.centroid_lst or self.measure_dist_lst):
            messagebox.showerror("Error Exporting Data","There is no data to export. \n\nPlease make sure to generate some data before exporting it.")
        else:
            workbook.close()
            self.progressbar_1['value']=100
            messagebox.showinfo("Export Successful",f'Your data is exported to the file at the following location. \n\n {file}')
            self.progressbar_1['value']=0
        
            
        

if __name__ == '__main__':
    import tkinter as tk
    root = tk.Tk()
    root.resizable(width=False, height=False)
    root.title("Point CLoud Application")
    #root.iconbitmap('logo_small.ico')
    app = NewprojectApp(root)
    app.run()


import tkinter as tk
import pandas as pd
import os


class Tk_Main(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(" ∫ ∬ ∭ ⨌∮∯∰ Spreadsheet / Dataframe customized tool ♬♩♩ ♬ ♩ ♬ ♬ ( work in progress.... ) Alex. ▛ ▜ ▙ ▟ ▚ ▞")
        self.geometry("800x600")
        self.rd_opt = 'label'
        
        self.r_dot, self.f_name = tk.StringVar(), tk.StringVar()
        self.ls_var, self.dict_rb = self.ls_dir_file(), {}
        
        self.f_m_bt = tk.Menubutton(text="  📨  📈  📉  Files to open 🚀  📝  📜", relief='ridge')
        self.rd_opt_bt = tk.Menubutton(text=" 📐   🔍   🔎  Render option 🔧   🌍   ⚡ ", relief='ridge')
        self.f_opt_bt = tk.Menubutton(text="  💾  🧪  💼  File option  🧩  📀  💽", relief='ridge')
        
        self.f_m = tk.Menu(self.f_m_bt, tearoff=0, relief='ridge')
        self.rd_opt_v = tk.Menu(self.rd_opt_bt, tearoff=0, relief='ridge')
        self.f_opt_v = tk.Menu(self.f_opt_bt, tearoff=0, relief='ridge')
        
        self.f_m_bt['menu'] = self.f_m
        self.rd_opt_bt['menu'] = self.rd_opt_v
        self.f_opt_bt['menu'] = self.f_opt_v
        
        self.f_m.add_command(label='List Excel', command=self.on_list_xl_clicked)
        
        self.rd_opt_v.add_command(label='Label', command=self.set_opt_lb)
        self.rd_opt_v.add_command(label='Text', command=self.set_opt_txt)
        
        self.f_opt_v.add_command(label='Save', command=self.save_c_to_xl)
        self.f_opt_v.add_command(label='Remove', command=self.remove_file)
        
        self.f_name_entry = tk.Entry(width=20, textvariable=self.f_name)
        self.f_m_bt.pack(fill='both')
        self.rd_opt_bt.pack(fill='both')
        self.f_opt_bt.pack(fill='both')
        self.f_name_entry.pack()     
    def open_xl(self, f_name):
        read_df = pd.read_excel('{}'.format(f_name), index_col=0)
        return read_df
    def save_xl(self, df_input, f_name):
        df_r = pd.DataFrame(df_input, index = None, columns = None)
        df_r.to_excel('{}.xlsx'.format(f_name), startrow=0, startcol=0)
    def ls_dir_file(self, directory=None, filetype='.xlsx'):
        ls_dir_raw = os.listdir(directory)
        ls_dir = []
        for i in ls_dir_raw:
            if filetype == 'any':
                ls_dir += [i]
            elif filetype in i:
                ls_dir += [i]
        return ls_dir
    def remove_file(self):
        read_df = os.remove(self.f_name.get())
    def on_list_xl_clicked(self):
        if len(self.ls_var) != len(self.dict_rb):
            for i in self.ls_var:
                self.dict_rb[i] = tk.Radiobutton(self, text=i, command=self.select_f,variable=self.r_dot, value=i)
                self.dict_rb[i].pack(anchor='w',fill='x', side='top')
    def set_opt_lb(self):
        self.rd_opt = 'label'
    def set_opt_txt(self):
        self.rd_opt = 'text' 
    def save_c_to_xl(self):
        self.save_xl(pd.DataFrame(eval(self.text_opt.get(1.0, tk.END))), self.f_name.get()+'.xlsx')
    def select_f(self):
        try: 
            self.lb_v.destroy()
            self.text_opt.destroy()
        except:
            pass
        self.f_name.set(self.r_dot.get().replace('.xlsx', ''))
        if self.rd_opt == 'label':
            self.lb_v = tk.Label(text=self.open_xl(self.r_dot.get()))
            self.lb_v.pack(fill='both',anchor='e', side='top')
        if self.rd_opt == 'text':
            self.text_opt = tk.Text(self, wrap="word", font=("Helvetica", 14))
            self.text_opt.insert(float(), self.open_xl(self.r_dot.get()).to_dict())
            self.text_opt.pack(fill='both', anchor='e',side='top')   
            
tk_init = Tk_Main()
tk_init.mainloop()
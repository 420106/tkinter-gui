from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from functools import partial
from threading import Thread
from datetime import datetime


def main():
    root = Tk()
    Application(root, config)
    root.mainloop()


class Application:

    def __init__(self, root, config):
        # Config
        root.title(config['title'])

        # Layout
        root.config(width=500)
        root.resizable(True, False)
        root.columnconfigure(1, weight=3)

        # Widgets
        ttk.Label(root, text=config['instruction']).grid(row=0, columnspan=3)

        self.str_vars = {}
        for i, f in enumerate(config['files']):
            ttk.Label(root, text=f) \
                .grid(row=i + 1, column=0)
            self.str_vars[i] = StringVar()
            ttk.Entry(root, state='disabled', textvariable=self.str_vars[i], width=50) \
                .grid(row=i + 1, column=1, sticky='nsew')
            ttk.Button(root, text='Browse', command=partial(self.browse_files, self.str_vars[i])) \
                .grid(row=i + 1, column=2)

        self.progress_bar = ttk.Progressbar(root, mode='indeterminate')
        self.progress_bar.grid(row=len(config['files']) + 1, column=0, columnspan=2, sticky='nsew')
        self.button_run = ttk.Button(root, text='Run', state='disabled', command=partial(self.thread, self.run_script))
        self.button_run.grid(row=len(config['files']) + 1, column=2)

        self.console = Text(root, state='disabled', height=10, background='DodgerBlue4', foreground='goldenrod')
        sys.stdout = TextRedirector(self.console)  # redirect print to GUI
        self.console.grid(row=len(config['files']) + 2, column=0, columnspan=3, sticky='nsew')
        self.scroll_bar = ttk.Scrollbar(root, orient=VERTICAL, command=self.console.yview)
        self.console.config(yscrollcommand=self.scroll_bar.set)
        self.scroll_bar.grid(row=len(config['files']) + 2, column=3, sticky='ns')

    def browse_files(self, target):
        '''Select a file and assign the file path to the target string variable'''
        filename = filedialog.askopenfilename(title='Browse', filetypes=(('All Files', '*.*'),
                                                                         ('Excel Files', '*.xls,*.xlsx')))
        target.set(filename)
        self.enable_run_button()

    def enable_run_button(self):
        '''Validate requested entries, if all filled, enable Run Button'''
        if all([v.get() for v in self.str_vars.values()]):
            self.button_run.configure(state='normal')

    @staticmethod
    def thread(func):
        '''Run a new thread for backend processing'''
        thread = Thread(target=func)
        thread.start()

    def run_script(self):
        self.progress_bar.start()
        get_timestamp = lambda: datetime.now().strftime("%d-%b-%Y %I:%M:%S %p")
        print(f'{get_timestamp()}: Start')
        try:
            ############### Load file path and do something here ###############
            param_1, param_2, param_3 = (v.get() for v in self.str_vars.values())


            ####################################################################
        except Exception as e:
            print(e)
        finally:
            print(f'{get_timestamp()}: End')
            self.progress_bar.stop()


class TextRedirector:

    def __init__(self, widget, tag='stdout'):
        self.widget = widget
        self.tag = tag

    def write(self, msg):
        self.widget.configure(state='normal')
        self.widget.insert(END, msg, self.tag)
        self.widget.configure(state='disabled')


config = {
    'title': '<application name>',
    'instruction': '<one-liner instruction>',
    'files': ['<input file 1>', '<input file 2>', '<input file 3>']
}

if __name__ == '__main__':
    main()

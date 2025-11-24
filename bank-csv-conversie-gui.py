#!/usr/bin/env python

#  Copyright (C) 2025 Toon van der Pas, Houten.
#
#  SPDX-License-Identifier: BSD-2-Clause
#
#  Redistribution and use in source and binary forms, with or without
#  modification, are permitted provided that the following conditions are met:
#
#  1. Redistributions of source code must retain the above copyright notice,
#     this list of conditions and the following disclaimer.
#
#  2. Redistributions in binary form must reproduce the above copyright notice,
#     this list of conditions and the following disclaimer in the documentation
#     and/or other materials provided with the distribution.
#
#  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS “AS IS” AND
#  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
#  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
#  IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT,
#  INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
#  BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
#  DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
#  LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE
#  OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED
#  OF THE POSSIBILITY OF SUCH DAMAGE.

import os.path
import sys
import array
from os.path import expanduser
import argparse
import configparser
import subprocess
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
import re, chardet, operator

class csv_conversie:
    def __init__(self, root):
        self.root = root
        self.title = "CSV-conversie"
        self.root.title(self.title)
        self.pad = 5
        self.grid_rownr_tree = 10 # row number of the tree frame
        self.dropdown_verbose_levels = [ "1 (ERROR)", "2 (WARN)", "3 (INFO)", "4 (DEBUG)" ]

        self.script     = os.path.realpath(__file__)
        self.script_dir = os.path.dirname(self.script)
        self.home_dir   = expanduser("~")

        self.var_script    = tk.StringVar(value = "")               # Should hold only the filename
        self.var_infile    = tk.StringVar(value = "")
        self.var_matchfile = tk.StringVar(value = "")  # Full path required
        self.var_outdir    = tk.StringVar(value = "")
        self.var_logfile   = tk.StringVar(value = "")
        self.var_verbosity = tk.StringVar(value = "2 (WARN)")
        if args.cfgfile == None:
            self.var_cfgfile   = tk.StringVar(value = "")
        else:
            self.var_cfgfile   = tk.StringVar(value=args.cfgfile)

        # Dit frame bevat de configuratie- en control-widgets.
        # Het biedt alléén horizontale resizing.
        frame_controls = tk.Frame(master=self.root, padx=10, pady=10)
        frame_controls.grid(row=0, column=0, sticky="ew")

        # Dit frame bevat de treeview waarin de mismatches uit de logfile worden getoond.
        # Het ondersteunt zowel horizontale als vertikale resizing.
        # De treeview is voorzien van horizontale en vertikale scrollbars.
        frame_treeview = tk.LabelFrame(master=self.root, padx=10, pady=10, text="Mutaties die niet gematched konden worden")
        frame_treeview.grid(row=self.grid_rownr_tree, column=0, padx=5, pady=5, sticky="nsew")

        # Maak de widgets aan in het controls-frame.
        self.lbl_script       = ttk.Label      (master=frame_controls, text="Script name")
        self.lbl_infile       = ttk.Label      (master=frame_controls, text="Infile")
        self.lbl_matchfile    = ttk.Label      (master=frame_controls, text="Matchfile")
        self.lbl_outdir       = ttk.Label      (master=frame_controls, text="Outdir")
        self.lbl_logfile      = ttk.Label      (master=frame_controls, text="Logfile")
        self.lbl_verbosity    = ttk.Label      (master=frame_controls, text="Verbosity")
        self.lbl_cfgfile      = ttk.Label      (master=frame_controls, text="Configfile")
        self.entry_script     = ttk.Entry      (master=frame_controls, textvariable=self.var_script)
        self.entry_infile     = ttk.Entry      (master=frame_controls, textvariable=self.var_infile)
        self.entry_matchfile  = ttk.Entry      (master=frame_controls, textvariable=self.var_matchfile)
        self.entry_outdir     = ttk.Entry      (master=frame_controls, textvariable=self.var_outdir)
        self.entry_logfile    = ttk.Entry      (master=frame_controls, textvariable=self.var_logfile)
        self.entry_cfgfile    = ttk.Entry      (master=frame_controls, textvariable=self.var_cfgfile)
        self.drpdwn_verbosity = ttk.OptionMenu        (frame_controls, self.var_verbosity,
                                                                       self.dropdown_verbose_levels[1],
                                                                      *self.dropdown_verbose_levels)
        self.bttn_select_infile    = ttk.Button (master=frame_controls, text="Select File", command=self.select_infile)
        self.bttn_select_matchfile = ttk.Button (master=frame_controls, text="Select File", command=self.select_matchfile)
        self.bttn_select_outdir    = ttk.Button (master=frame_controls, text="Select Dir",  command=self.select_outdir)
        self.bttn_select_cfgfile   = ttk.Button (master=frame_controls, text="Select File", command=self.select_cfgfile)

        self.bttn_load_cfg = ttk.Button (master=frame_controls, text="Load Config", command=self.load_config)
        self.bttn_save_cfg = ttk.Button (master=frame_controls, text="Save Config", command=self.save_config)
        self.bttn_run_csv_conversion = ttk.Button (master=frame_controls, text="Run CSV Conversion", command=self.run_csv_conversion_script)

        self.sep_top       = ttk.Separator (master=frame_controls, orient=tk.HORIZONTAL)
        self.sep_bottom    = ttk.Separator (master=frame_controls, orient=tk.HORIZONTAL)

        # Maak de widgets aan in het treeview-frame.
        self.hsb = ttk.Scrollbar(master=frame_treeview, orient="horizontal")   # Horizontale scrollbar
        self.vsb = ttk.Scrollbar(master=frame_treeview, orient="vertical")     # Vertikale scrollbar
        self.tree = ttk.Treeview(master=frame_treeview, columns=("bedrag", "tegenrekening", "tegenpartij", "regarding"),
                                                        show="headings",
                                                        xscrollcommand=self.hsb.set,
                                                        yscrollcommand=self.vsb.set)
        self.hsb.config(command=self.tree.xview)
        self.vsb.config(command=self.tree.yview)

        self.tree.heading("bedrag",        text="Bedrag")
        self.tree.heading("tegenrekening", text="Tegenrekening")
        self.tree.heading("tegenpartij",   text="Tegenpartij")
        self.tree.heading("regarding",     text="Regarding")
        self.tree.column('bedrag',        anchor='e', width=200, minwidth=75)
        self.tree.column('tegenrekening', anchor='w', width=200, minwidth=165)
        self.tree.column('tegenpartij',   anchor='w', width=200, minwidth=260)
        self.tree.column('regarding',     anchor='w', width=200, minwidth=1900)

        # Plaats de widgets in het grid.
        self.lbl_script.grid              (row=0, column=0, padx=self.pad, pady=self.pad, sticky="w")
        self.entry_script.grid            (row=0, column=1, padx=self.pad, pady=self.pad, sticky=tk.EW, columnspan=2)
        self.sep_top.grid                 (row=1, column=0, padx=self.pad, pady=self.pad, sticky=tk.EW, columnspan=4)
        self.lbl_infile.grid              (row=2, column=0, padx=self.pad, pady=self.pad, sticky="w")
        self.entry_infile.grid            (row=2, column=1, padx=self.pad, pady=self.pad, sticky=tk.EW, columnspan=2)
        self.bttn_select_infile.grid      (row=2, column=3, padx=self.pad, pady=self.pad)
        self.lbl_matchfile.grid           (row=3, column=0, padx=self.pad, pady=self.pad, sticky="w")
        self.entry_matchfile.grid         (row=3, column=1, padx=self.pad, pady=self.pad, sticky=tk.EW, columnspan=2)
        self.bttn_select_matchfile.grid   (row=3, column=3, padx=self.pad, pady=self.pad)
        self.lbl_outdir.grid              (row=4, column=0, padx=self.pad, pady=self.pad, sticky="w")
        self.entry_outdir.grid            (row=4, column=1, padx=self.pad, pady=self.pad, sticky=tk.EW, columnspan=2)
        self.bttn_select_outdir.grid      (row=4, column=3, padx=self.pad, pady=self.pad)
        self.lbl_logfile.grid             (row=5, column=0, padx=self.pad, pady=self.pad, sticky="w")
        self.entry_logfile.grid           (row=5, column=1, padx=self.pad, pady=self.pad, sticky=tk.EW, columnspan=2)
        self.lbl_verbosity.grid           (row=6, column=0, padx=self.pad, pady=self.pad, sticky="w")
        self.drpdwn_verbosity.grid        (row=6, column=1, padx=self.pad, pady=self.pad, sticky=tk.W,  columnspan=2)
        self.sep_bottom.grid              (row=7, column=0, padx=self.pad, pady=self.pad, sticky=tk.EW, columnspan=4)
        self.lbl_cfgfile.grid             (row=8, column=0, padx=self.pad, pady=self.pad, sticky="w")
        self.entry_cfgfile.grid           (row=8, column=1, padx=self.pad, pady=self.pad, sticky=tk.EW, columnspan=2)
        self.bttn_select_cfgfile.grid     (row=8, column=3, padx=self.pad, pady=self.pad)

        self.bttn_load_cfg.grid           (row=9, column=0, padx=self.pad, pady=self.pad, sticky=tk.W)
        self.bttn_save_cfg.grid           (row=9, column=1, padx=self.pad, pady=self.pad, sticky=tk.W)
        self.bttn_run_csv_conversion.grid (row=9, column=3, padx=self.pad, pady=self.pad, sticky=tk.W)

        self.tree.grid (row=self.grid_rownr_tree, column=0, padx=self.pad, pady=self.pad, sticky=tk.NSEW)
        self.vsb.grid  (row=self.grid_rownr_tree, column=1, sticky=tk.NS)
        self.hsb.grid  (row=self.grid_rownr_tree+1, column=0, sticky=tk.EW)

        # Configureer het resizing gedrag.
        self.root.grid_rowconfigure(self.grid_rownr_tree, weight=1) # Het root-window wat de frames bevat dient te resizen,
        self.root.grid_columnconfigure(0, weight=1)                 # want anders zullen de frames niet kunnen resizen.

        frame_controls.grid_columnconfigure(1, weight=1)            # Laat het controls-frame horizontaal resizen.

        frame_treeview.grid_rowconfigure(self.grid_rownr_tree, weight=1) # Laat het treeview-frame vertikaal resizen.
        frame_treeview.grid_columnconfigure(0, weight=1)                 # Laat het treeview-frame horizontaal resizen.

        # Some silly manipulations to make us cross-compatible with the Windows filesystem.
    def fix_path(self, path):
        if os.altsep is not None:
            path = path.replace(os.altsep, os.sep)
        if os.name=='nt':
            # We are on Windows.
            # If path doesn't contain a drive character,
            # add the one from the home directory we determined earlier.
            if not re.search('^[A-Z]:', path):
                path=self.home_dir[:2] + path
        else:
            # We are not on Windows but (hopefully) a POSIX system.
            # Strip de drive character if it exists.
            if re.search('^[A-Z]:', path):
                path=path[2:]
        return path

    def set_entry(self, entry, text):
        entry.delete(0, tk.END)
        entry.insert(0, text)

    def select_infile(self):
        tmp_var = ""
        init_dir = self.entry_infile.get()
        if init_dir == "":
            init_dir = self.home_dir
        filetypes = (
            ('CSV-infiles', '*.csv'),
            ('All files', '*.*')
        )

        tmp_var = self.fix_path(fd.askopenfilename(
            initialdir=init_dir,
            title='Select an export CSV-file from your bank account',
            filetypes=filetypes))
        self.set_entry(self.entry_infile, tmp_var)

    def select_matchfile(self):
        tmp_var = ""
        init_dir = self.entry_matchfile.get()
        if init_dir == "":
            init_dir = self.home_dir
        filetypes = (
            ('CSV-matchfiles', '*.csv'),
            ('All files', '*.*')
        )

        tmp_var = self.fix_path(fd.askopenfilename(
            initialdir=init_dir,
            title='Select a CSV file containing your matching rules',
            filetypes=filetypes))
        self.set_entry(self.entry_matchfile, tmp_var)

    def select_outdir(self):
        init_dir = self.entry_outdir.get()
        if init_dir == "":
            init_dir = self.home_dir
        tmp_var = self.fix_path(fd.askdirectory(
            initialdir=init_dir,
            title='Select a directory for the output files'))
        self.set_entry(self.entry_outdir, tmp_var)

    def select_cfgfile(self):
        tmp_var  = ""
        init_dir = os.path.dirname(self.entry_cfgfile.get())
        if init_dir == "":
            init_dir = self.home_dir
        filetypes = (
            ('Log files', '*.ini'),
            ('All files', '*.*')
        )

        tmp_var = self.fix_path(fd.askopenfilename(
            initialdir=init_dir,
            title='Select an existing ini file',
            filetypes=filetypes))
        self.set_entry(self.entry_cfgfile, tmp_var)

    def save_config(self):
        config = configparser.ConfigParser()
        config['CSV-conversion'] = {'script':    self.var_script.get(),
                                    'infile':    self.var_infile.get(),
                                    'matchfile': self.var_matchfile.get(),
                                    'outdir':    self.var_outdir.get(),
                                    'logfile':   self.var_logfile.get(),
                                    'verbosity': self.var_verbosity.get(),
                                    'cfgfile':   self.var_cfgfile.get()}
        with open(self.var_cfgfile.get(), 'w') as configfile:
            config.write(configfile)

    def load_config(self):
        config = configparser.ConfigParser()
        config.read(self.var_cfgfile.get())
        self.set_entry(self.entry_script,                  config['CSV-conversion']['script'])
        self.set_entry(self.entry_infile,    self.fix_path(config['CSV-conversion']['infile']))
        self.set_entry(self.entry_matchfile, self.fix_path(config['CSV-conversion']['matchfile']))
        self.set_entry(self.entry_outdir,    self.fix_path(config['CSV-conversion']['outdir']))
        self.set_entry(self.entry_logfile,   self.fix_path(config['CSV-conversion']['logfile']))
        self.var_verbosity.set(              config['CSV-conversion']['verbosity'])
        self.set_entry(self.entry_cfgfile,   self.fix_path(config['CSV-conversion']['cfgfile']))

    def create_text_window(self, severity, text):
        text_win=tk.Toplevel(root)
        text_win.geometry("400x250")
        text_win.title(severity)
        frame_text_box = tk.Frame(text_win, padx=10, pady=10)
        frame_text_box.grid(row=0, column=0, sticky="nsew")
        frame_buttons = tk.Frame(text_win, padx=10, pady=10)
        frame_buttons.grid(row=1, column=0, sticky="ew")
        text_box  = tk.Text(frame_text_box, relief=tk.SOLID)
        button    = ttk.Button    (frame_buttons,  text="Close",      command=text_win.destroy)
        scrollbar = ttk.Scrollbar (frame_text_box, orient="vertical", command=text_box.yview)
        text_box['yscrollcommand'] = scrollbar.set

        text_box.insert(tk.END, text)

        text_box.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        scrollbar.grid(row=0, column=1, sticky="ns")
        button.grid(row=1, column=0, padx=5, pady=5, columnspan=2)

        text_win.grid_rowconfigure(0, weight=1)
        text_win.grid_columnconfigure(0, weight=1)
        frame_text_box.grid_rowconfigure(0, weight=1)
        frame_text_box.grid_columnconfigure(0, weight=1)
        frame_buttons.grid_columnconfigure(0, weight=1)

        self.root.wait_window(text_win)

    def parse_logfile_into_tree_window(self):
        unmatched_record_dict_list=[]

        try:
            with open(self.var_logfile.get(), 'rb') as rawdata:
                chardet_result = chardet.detect(rawdata.read(100000))
                logfile_encoding = chardet_result['encoding']
                rawdata.close()
        except (OSError, IOError) as e:
            self.create_text_window("ERROR", "Failed to read rawdata from logfile:\n" + str(e))
            sys.exit(2)

        try:
            with open(self.var_logfile.get(), 'r', encoding=logfile_encoding) as logfile:
                for line in logfile:
                    # We are only interested in records concerning failed matches...
                    if re.search('[0-9]{14} WARN No match on row number', line):
                        bedrag_start = re.search(' bedrag:\\s+', line).span()[1]
                        bedrag_end   = re.search(',', line[bedrag_start:]).span()[0]+bedrag_start+3
                        bedrag       = line[bedrag_start:bedrag_end]
                        tegenrekening_start = re.search('tegenrekening: ', line).span()[1]
                        tegenrekening_end   = re.search(',', line[tegenrekening_start:]).span()[0]+tegenrekening_start
                        tegenrekening       = line[tegenrekening_start:tegenrekening_end]
                        tegenpartij_start = re.search('tegenpartij: \'', line).span()[1]
                        tegenpartij_end   = re.search('\'', line[tegenpartij_start:]).span()[0]+tegenpartij_start
                        tegenpartij       = line[tegenpartij_start:tegenpartij_end]
                        regarding_start = re.search('regarding: \'', line).span()[1]
                        regarding_end   = re.search('\'', line[regarding_start:]).span()[0]+regarding_start
                        regarding       = line[regarding_start:regarding_end]
                        unmatched_record_dict_list.append({
                            'Bedrag':        bedrag,
                            'Tegenrekening': tegenrekening,
                            'Tegenpartij':   tegenpartij,
                            'Regarding':     regarding
                        })

                unmatched_record_dict_list_sorted=sorted(unmatched_record_dict_list,
                                                         key=operator.itemgetter('Tegenpartij',
                                                                                 'Tegenrekening'))

                for row in unmatched_record_dict_list_sorted:
                    self.tree.insert("", tk.END, values=(row['Bedrag'],
                                                         row['Tegenrekening'],
                                                         row['Tegenpartij'],
                                                         row['Regarding']))
        except (OSError, IOError) as e:
            self.create_text_window("ERROR", "Failed to read mismatches from logfile:\n" + str(e))
            sys.exit(2)

    def run_csv_conversion_script(self):
        arguments=[]
        arguments.append(sys.executable)
        arguments.append(self.script_dir + "/" + self.var_script.get())
        arguments.append("--infile")
        arguments.append(self.var_infile.get())
        arguments.append("--matchfile")
        arguments.append(self.var_matchfile.get())
        if self.entry_outdir.get() != "":
            arguments.append("--outdir")
            arguments.append(self.var_outdir.get())
        if self.entry_logfile.get() != "":
            arguments.append("--logfile")
            arguments.append(self.var_logfile.get())
        arguments.append("--verbosity")
        arguments.append(self.var_verbosity.get()[:1])

        try:
            if os.path.exists(self.var_logfile.get()):
                os.remove(self.var_logfile.get())
        except (OSError, IOError) as e:
            self.create_text_window("ERROR", "Failed to delete logfile:\n" + str(e))
            sys.exit(2)

        try:
            proc = subprocess.run(arguments, encoding='utf-8', check=True, capture_output=True)
            self.parse_logfile_into_tree_window()
            self.create_text_window("INFO", "CSV conversion successful")
        except subprocess.CalledProcessError as e:
            message="CSV conversion failed with exit code " + str(e.returncode) +  "!\n"
            if e.stdout is not None:
                message=message + "stdout:\n" + e.stdout
            if e.stderr is not None:
                message=message + "stderr:\n" + e.stderr
            message=message + "\n" + "The command was:\n"
            for arg in e.cmd:
                message=message + arg + " "
            self.create_text_window("ERROR", message)
            sys.exit(2)


if __name__ == "__main__":

    parser = argparse.ArgumentParser(description='This program provides a GUI for the rabo-csv.py and ing-csv.py programs.')
    parser.add_argument('--cfgfile', default=None)
    args = parser.parse_args()

    root = tk.Tk()
    app = csv_conversie(root)
    root.mainloop()

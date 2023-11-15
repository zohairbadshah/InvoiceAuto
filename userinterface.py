import subprocess
import tkinter as tk
from tkinter import filedialog
def browse_main_file():
    file_path = filedialog.askopenfilename(
        title="Select main file",
        filetypes=(("Excel File", "*.xls"), ("All files", "*.*")),

        initialdir="/Projects/WebAutoMindful-1/Variables")
    selected_files["Main file"] = file_path
    main_file.set(file_path)


def browse_accessories_file():
    file_path = filedialog.askopenfilename(
        title="Select accessories file",
        filetypes=(("Csv File", "*.csv"), ("All files", "*.*")),

        initialdir="/Projects/WebAutoMindful-1/Variables")
    selected_files["Accessories file"] = file_path
    accessories_file.set(file_path)


def browse_yarn_file():
    file_path = filedialog.askopenfilename(
        title="Select yarn file",
        filetypes=(("Csv File", "*.csv"), ("All files", "*.*")),

        initialdir="/Projects/WebAutoMindful-1/Variables")  # You can change the initial directory as needed
    selected_files["Yarn file"] = file_path
    yarn_file.set(file_path)


def browse_cost_center_code_file():
    file_path = filedialog.askopenfilename(
        title="Select cost center code file",
        filetypes=(("Excel File", "*.xlsx"), ("All files", "*.*")),

        initialdir="/Projects/WebAutoMindful-1/Variables")
    selected_files["Cost center code"] = file_path
    cost_center_code_file.set(file_path)


def browse_applicant_name_id_file():
    file_path = filedialog.askopenfilename(
        title="Select applicant name ID file",
        filetypes=(("Excel File", "*.xlsx"), ("All files", "*.*")),

        initialdir="/Projects/WebAutoMindful-1/Variables")
    selected_files["Applicant name Id"] = file_path
    applicant_name_id_file.set(file_path)

def submit():
    print(selected_files)
    subprocess.run(["python", "matchdata.py", str(selected_files)])
    root.destroy()
    input("Press Enter to exit")
if __name__ == "__main__":
    root = tk.Tk()
    root.title("File Selector")

    main_file = tk.StringVar()
    accessories_file = tk.StringVar()
    yarn_file = tk.StringVar()
    cost_center_code_file = tk.StringVar()
    applicant_name_id_file = tk.StringVar()

    selected_files = {
        "Main file": "",
        "Accessories file": "",
        "Yarn file": "",
        "Cost center code": "",
        "Applicant name Id": ""
    }





    # Create buttons to browse files
    main_button = tk.Button(root, text="Select main file", command=browse_main_file)
    accessories_button = tk.Button(root, text="Select accessories file", command=browse_accessories_file)
    yarn_button = tk.Button(root, text="Select yarn file", command=browse_yarn_file)
    cost_center_code_button = tk.Button(root, text="Select cost center code file", command=browse_cost_center_code_file)
    applicant_name_id_button = tk.Button(root, text="Select applicant name ID file",
                                         command=browse_applicant_name_id_file)
    submit_button = tk.Button(root, text="Submit", command=submit)

    # Create labels to display selected files
    main_label = tk.Label(root, textvariable=main_file, justify="left")
    accessories_label = tk.Label(root, textvariable=accessories_file, justify="left")
    yarn_label = tk.Label(root, textvariable=yarn_file, justify="left")
    cost_center_code_label = tk.Label(root, textvariable=cost_center_code_file, justify="left")
    applicant_name_id_label = tk.Label(root, textvariable=applicant_name_id_file, justify="left")

    # Pack buttons and labels
    main_button.pack()
    main_label.pack()
    accessories_button.pack()
    accessories_label.pack()
    yarn_button.pack()
    yarn_label.pack()
    cost_center_code_button.pack()
    cost_center_code_label.pack()
    applicant_name_id_button.pack()
    applicant_name_id_label.pack()
    submit_button.pack()
    root.mainloop()


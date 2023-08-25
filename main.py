import customtkinter as moderntk
import pyodbc

DB_FILE_PATH = r'G:\Github\Command-Explorer\db\DataBackend.accdb'
TABLE_NAME = 'CommandData'

class CommandExplorerApp:
    def __init__(self):
        self.app = moderntk.CTk()
        self.app.geometry("720x480")
        self.app.title("ToolKit")
        self.app.resizable(False, False)
        self.appearance = moderntk.set_appearance_mode("dark")
        self.setup_ui()

    def setup_ui(self):
        self.setup_fonts()
        self.create_widgets()
        self.populate_command_combo()

    def setup_fonts(self):
        self.text_font = ("Helvetica", 14)
        self.result_font = ("Courier", 14)
        self.header_font = moderntk.CTkFont(family="Franklin Gothic Demi", size=30, weight="bold", slant="italic")
        self.header_font_v2 = ("Helvetica", 16)
        self.footer_font = moderntk.CTkFont(family="Helvetica", size=12, slant="italic")

    def create_widgets(self):
        self.message_label = moderntk.CTkLabel(self.app, text="Command Explorer", font=self.header_font, fg_color="transparent")
        self.message_label.grid(row=0, column=0, padx=20, pady=20)

        self.command_radio_flag = moderntk.IntVar(value=0)
        self.command_radio_no1 = moderntk.CTkRadioButton(self.app, text="Search Mode", command=self.radiobutton_event, variable=self.command_radio_flag, value=1)
        self.command_radio_no2 = moderntk.CTkRadioButton(self.app, text="Select Mode", command=self.radiobutton_event, variable=self.command_radio_flag, value=2)
        self.command_radio_no1.grid(row=1, column=0, padx=20, pady=15, sticky="w")
        self.command_radio_no2.grid(row=1, column=0, padx=150, pady=15, sticky="w")

        self.command = moderntk.CTkEntry(self.app, width=30, height=30, placeholder_text="Enter the Command", font=self.text_font)
        #self.command.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        #self.command.bind("<Return>", self.generate_results)

        self.command_combo = moderntk.CTkOptionMenu(self.app, width=30, height=30, font=self.text_font, command=self.generate_checkbox_results, fg_color="GREY20")
        #self.command_combo.grid(row=3, column=0, padx=20, pady=(10, 10), sticky="w")
        #self.command_combo.configure(width=170)

        self.button = moderntk.CTkButton(self.app, width=70, height=20, text="Search", command=self.generate_results)
        #self.button.grid(row=4, column=0, padx=20, pady=20, sticky="ew")

        self.message_label_v2 = moderntk.CTkLabel(self.app, text="Syntax:", font=self.header_font_v2)
        self.message_label_v2.grid(row=5, column=0, padx=20, pady=2, sticky="w")

        self.syntax = moderntk.CTkTextbox(self.app, width=40, height=40, state=moderntk.NORMAL, font=self.result_font)
        self.syntax.grid(row=6, column=0, padx=20, pady=10, sticky="nsew")

        self.message_label_v3 = moderntk.CTkLabel(self.app, text="Example:", font=self.header_font_v2)
        self.message_label_v3.grid(row=7, column=0, padx=20, pady=2, sticky="w")

        self.syntax_ex = moderntk.CTkTextbox(self.app, width=40, height=40, state=moderntk.NORMAL, font=self.result_font)
        self.syntax_ex.grid(row=9, column=0, padx=20, pady=10, sticky="nsew")

        self.credits = moderntk.CTkLabel(self.app, text="Made with ♥ by Anush", font=self.footer_font, fg_color="transparent")
        self.credits.grid(row=10, column=0, padx=5, pady=5, sticky="")

        self.app.grid_rowconfigure(0, weight=1)
        self.app.grid_columnconfigure(0, weight=1)

    def connect_to_access(self):
        try:
            conn = pyodbc.connect(f'Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={DB_FILE_PATH};')
            return conn
        except Exception as e:
            self.handle_error(f"Error connecting to MS Access: {e}")
            return None

    def fetch_data_from_db(self, conn, command_name):
        cursor = conn.cursor()
        cursor.execute(f"select * from {TABLE_NAME} where command=?", (command_name,))
        return cursor.fetchall()

    def handle_error(self, error_message):
        self.message_label.configure(text=error_message)
        print(error_message)

    def populate_command_combo(self):
        conn = self.connect_to_access()
        if conn:
            cursor = conn.cursor()
            cursor.execute(f"SELECT DISTINCT command FROM {TABLE_NAME}")
            commands = [row[0] for row in cursor.fetchall()]
            self.command_combo.configure(values=commands)
            self.command_combo.set("Select to Search")
            conn.close()

    def generate_results(self, event=0):
        command_name = self.command.get()
        self.update_syntax_views(command_name)

    def generate_checkbox_results(self, choice):
        self.update_syntax_views(choice)
    
    def radiobutton_event(self, event=0):
        mode = self.command_radio_flag.get()
        if mode == 1:
            #self.command = moderntk.CTkEntry(self.app, width=40, height=40, placeholder_text="Enter the Command", font=self.text_font)
            self.command.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
            self.command.bind("<Return>", self.generate_results)

            #self.button = moderntk.CTkButton(self.app, text="Search", command=self.generate_results)
            self.button.grid(row=2, column=0, padx=25, pady=0, sticky="e")

            self.command_combo.grid_remove()

        elif mode == 2:
            self.command.grid_remove()
            self.button.grid_remove()

            #self.command_combo = moderntk.CTkOptionMenu(self.app, font=self.text_font, command=self.generate_checkbox_results, fg_color="GREY20")
            self.command_combo.grid(row=3, column=0, padx=20, pady=(10, 10), sticky="w")
            self.command_combo.configure(width=170)

    def update_syntax_views(self, command_name):
        self.syntax.configure(state=moderntk.NORMAL)
        self.syntax_ex.configure(state=moderntk.NORMAL)
        self.syntax.delete("1.0", "end")
        self.syntax_ex.delete("1.0", "end")

        conn = self.connect_to_access()
        if conn:
            rows = self.fetch_data_from_db(conn, command_name)
            for row in rows:
                self.syntax.insert("0.0", str(row[1]))
                self.syntax_ex.insert("0.0", str(row[2]))
            self.syntax.configure(state=moderntk.DISABLED)
            self.syntax_ex.configure(state=moderntk.DISABLED)
            conn.close()

    def run(self):
        self.app.mainloop()

if __name__ == "__main__":
    app = CommandExplorerApp()
    app.run()
import customtkinter as ctk
from tkinter import messagebox
from PIL import Image, ImageTk
import hashlib
from utils import resource_path

COLORS = {
    "primary": "#1f6aa5",
    "secondary": "#2b2b2b",
    "success": "#2CC985",
    "error": "#E74C3C",
    "warning": "#F39C12",
    "info": "#3498DB",
    "background": "#1e1e1e",
    "card": "#2b2b2b",
    "text": "#ffffff",
    "text_secondary": "#a0a0a0"
}

DEFAULT_USERS = {
    "manager": hashlib.sha256("manager123".encode()).hexdigest()
}


class LoginFrame:
    def __init__(self, parent, on_success_callback):
        self.parent = parent
        self.on_success_callback = on_success_callback

        # Create main container frame
        self.main_container = ctk.CTkFrame(parent, fg_color=COLORS["background"], corner_radius=0)
        self.main_container.pack(fill='both', expand=True)

        self.setup_ui()

        # Bind Enter key to login
        parent.bind('<Return>', lambda e: self.login())

    def setup_ui(self):
        # Content frame 
        content_frame = ctk.CTkFrame(self.main_container, fg_color=COLORS["background"])
        content_frame.pack(fill='both', expand=True, padx=0, pady=0)

        # Left side 
        screen_width = self.parent.winfo_screenwidth()
        left_width = int(screen_width * 0.4)
        left_frame = ctk.CTkFrame(content_frame, fg_color=COLORS["primary"], corner_radius=0, width=left_width)
        left_frame.pack(side='left', fill='both', expand=True, padx=0, pady=0)
        left_frame.pack_propagate(False)

        # Logo and branding
        branding_container = ctk.CTkFrame(left_frame, fg_color="transparent")
        branding_container.pack(expand=True, pady=50)

        # Try to load logo
        try:
            img = Image.open(resource_path("mbs_logo.png"))
            img = img.resize((200, 200), Image.Resampling.LANCZOS)
            ctk_image = ctk.CTkImage(light_image=img, dark_image=img, size=(200, 200))
            logo_label = ctk.CTkLabel(branding_container, image=ctk_image, text="")
            logo_label.pack(pady=30)
        except:
            logo_label = ctk.CTkLabel(branding_container, text="🎓",
                                      font=("Arial", 120), text_color="white")
            logo_label.pack(pady=30)

        welcome_label = ctk.CTkLabel(branding_container,
                                     text="TRAINING & PLACEMENT\nMANAGEMENT SYSTEM",
                                     font=("Arial", 32, "bold"),
                                     text_color="white",
                                     justify="center")
        welcome_label.pack(pady=20)

        subtitle_label = ctk.CTkLabel(branding_container,
                                      text="Mahant Bachittar Singh College\nof Engineering & Technology",
                                      font=("Arial", 18),
                                      text_color="white",
                                      justify="center")
        subtitle_label.pack(pady=15)

        # Right side - Login form
        right_frame = ctk.CTkFrame(content_frame, fg_color=COLORS["background"], corner_radius=0)
        right_frame.pack(side='right', fill='both', expand=True, padx=0, pady=0)

        # Login form container
        form_container = ctk.CTkFrame(right_frame, fg_color="transparent")
        form_container.pack(expand=True, pady=50)

        # Login header
        login_header = ctk.CTkLabel(form_container, text="Welcome Back!",
                                    font=("Arial", 42, "bold"),
                                    text_color=COLORS["text"])
        login_header.pack(pady=(0, 15))

        login_subheader = ctk.CTkLabel(form_container, text="Please login to continue",
                                       font=("Arial", 18),
                                       text_color=COLORS["text_secondary"])
        login_subheader.pack(pady=(0, 50))

        # Username field
        username_label = ctk.CTkLabel(form_container, text="Username",
                                      font=("Arial", 16, "bold"),
                                      text_color=COLORS["text"],
                                      anchor="w")
        username_label.pack(fill='x', padx=80, pady=(0, 8))

        self.username_entry = ctk.CTkEntry(form_container,
                                           placeholder_text="Enter your username",
                                           width=450,
                                           height=55,
                                           font=("Arial", 16),
                                           border_width=2,
                                           corner_radius=10)
        self.username_entry.pack(padx=80, pady=(0, 25))
        self.username_entry.focus()

        # Password field
        password_label = ctk.CTkLabel(form_container, text="Password",
                                      font=("Arial", 16, "bold"),
                                      text_color=COLORS["text"],
                                      anchor="w")
        password_label.pack(fill='x', padx=80, pady=(0, 8))

        self.password_entry = ctk.CTkEntry(form_container,
                                           placeholder_text="Enter your password",
                                           show="●",
                                           width=450,
                                           height=55,
                                           font=("Arial", 16),
                                           border_width=2,
                                           corner_radius=10)
        self.password_entry.pack(padx=80, pady=(0, 15))

        # Show/Hide password checkbox
        self.show_password_var = ctk.BooleanVar(value=False)
        show_password_check = ctk.CTkCheckBox(form_container,
                                              text="Show Password",
                                              variable=self.show_password_var,
                                              command=self.toggle_password,
                                              font=("Arial", 14),
                                              text_color=COLORS["text_secondary"])
        show_password_check.pack(pady=(0, 35))

        # Login button
        login_btn = ctk.CTkButton(form_container,
                                  text="LOGIN",
                                  command=self.login,
                                  width=450,
                                  height=60,
                                  font=("Arial", 18, "bold"),
                                  fg_color=COLORS["primary"],
                                  hover_color=COLORS["info"],
                                  corner_radius=10)
        login_btn.pack(pady=(0, 25))

        # Info section
        info_frame = ctk.CTkFrame(form_container, fg_color=COLORS["card"], corner_radius=10)
        info_frame.pack(fill='x', padx=80, pady=(25, 0))

        info_label = ctk.CTkLabel(info_frame,
                                  text="please enter Username and Password",
                                  font=("Arial", 14),
                                  text_color=COLORS["text_secondary"],
                                  justify="left")
        info_label.pack(padx=25, pady=20)

        # Footer
        footer_label = ctk.CTkLabel(right_frame,
                                    text="© 2024 MBS College - All Rights Reserved",
                                    font=("Arial", 12),
                                    text_color=COLORS["text_secondary"])
        footer_label.pack(side='bottom', pady=25)

    def toggle_password(self):
        """Toggle password visibility"""
        if self.show_password_var.get():
            self.password_entry.configure(show="")
        else:
            self.password_entry.configure(show="●")

    def login(self):
        """Handle login authentication"""
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()

        if not username or not password:
            messagebox.showerror("Login Failed", "Please enter both username and password!")
            return

        # Hash the entered password
        hashed_password = hashlib.sha256(password.encode()).hexdigest()

        # Check credentials
        if username in DEFAULT_USERS and DEFAULT_USERS[username] == hashed_password:
            messagebox.showinfo("Login Successful", f"Welcome, {username}!")
            self.main_container.destroy()
            self.on_success_callback(username)
        else:
            messagebox.showerror("Login Failed",
                                 "Invalid username or password!\n\nPlease check your credentials and try again.")
            self.password_entry.delete(0, 'end')
            self.password_entry.focus()



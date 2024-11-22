import tkinter as tk
from tkinter import messagebox
import openpyxl
from openpyxl import Workbook
import os

class QuizApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Python Quiz App")
        self.root.geometry("500x500")
        self.root.config(bg="#6C7B8B")

        self.questions = [
            {
                "question": "What is the output of 3 + 2 * 2?",
                "options": ["7", "10", "5", "8"],
                "correct": "7"
            },
            {
                "question": "What is the keyword to define a function in Python?",
                "options": ["function", "def", "func", "define"],
                "correct": "def"
            },
            {
                "question": "Which of the following is a mutable data type?",
                "options": ["Tuple", "List", "String", "Set"],
                "correct": "List"
            },
            {
                "question": "How do you start a comment in Python?",
                "options": ["//", "#", "/*", "--"],
                "correct": "#"
            }
        ]

        self.question_index = 0
        self.correct_answers = 0
        self.total_marks = len(self.questions)
        self.time_limit = 60  # Time limit in seconds

        self.name = ""
        self.sapid = ""

        self.create_user_info_widgets()

    def create_user_info_widgets(self):
        """Create widgets for entering user details."""
        self.title_label = tk.Label(self.root, text="Enter Your Details", font=("Verdana", 16, "bold"), bg="#6C7B8B", fg="#FFF")
        self.title_label.pack(pady=20)

        tk.Label(self.root, text="Name:", font=("Arial", 12), bg="#6C7B8B", fg="#FFF").pack(pady=5)
        self.name_entry = tk.Entry(self.root, font=("Arial", 12))
        self.name_entry.pack(pady=5)

        tk.Label(self.root, text="SAP ID:", font=("Arial", 12), bg="#6C7B8B", fg="#FFF").pack(pady=5)
        self.sapid_entry = tk.Entry(self.root, font=("Arial", 12))
        self.sapid_entry.pack(pady=5)

        self.start_button = tk.Button(self.root, text="Start Quiz", font=("Arial", 14), bg="#4CAF50", fg="white",
                                      command=self.start_quiz)
        self.start_button.pack(pady=20)

    def start_quiz(self):
        """Start the quiz after collecting user details."""
        self.name = self.name_entry.get().strip()
        self.sapid = self.sapid_entry.get().strip()

        if not self.name or not self.sapid:
            messagebox.showerror("Error", "Please fill in all fields!")
        else:
            # Clear user input widgets
            self.title_label.destroy()
            self.name_entry.destroy()
            self.sapid_entry.destroy()
            self.start_button.destroy()

            # Create quiz widgets
            self.create_quiz_widgets()
            self.start_timer()
            self.load_question()

    def create_quiz_widgets(self):
        """Create widgets for the quiz."""
        self.title_label = tk.Label(self.root, text=f"Python Quiz - {self.name}", font=('Verdana', 20, 'bold'), fg="#FFF", bg="#6C7B8B")
        self.title_label.pack(pady=20)

        self.timer_label = tk.Label(self.root, text=f"Time Left: {self.time_limit}s", font=('Arial', 14), bg="#6C7B8B", fg="#FFF")
        self.timer_label.pack(pady=5)

        self.question_label = tk.Label(self.root, text="", font=('Arial', 14), wraplength=350, bg="#6C7B8B", fg="#FFF")
        self.question_label.pack(pady=20)

        self.option_buttons = []
        for i in range(4):
            btn = tk.Button(self.root, text="", width=20, height=2, font=('Arial', 12), command=lambda i=i: self.check_answer(i),
                            bd=3, bg="#4CAF50", fg="white", activebackground="#45a049", activeforeground="white", 
                            relief="flat", cursor="hand2", highlightthickness=0)
            btn.pack(pady=10, padx=20)
            self.option_buttons.append(btn)

        self.next_button = tk.Button(self.root, text="Next", width=10, height=2, font=('Arial', 14), command=self.load_next_question,
                                     bd=3, bg="#FF9800", fg="white", activebackground="#e68917", activeforeground="white", 
                                     relief="flat", cursor="hand2", highlightthickness=0, state="disabled")
        self.next_button.pack(pady=15)

        self.score_label = tk.Label(self.root, text=f"Score: {self.correct_answers}/{self.total_marks}", font=('Arial', 12), bg="#6C7B8B", fg="#FFF")
        self.score_label.pack(pady=10)

    def start_timer(self):
        """Start the countdown timer."""
        if self.time_limit > 0:
            self.timer_label.config(text=f"Time Left: {self.time_limit}s")
            self.time_limit -= 1
            self.root.after(1000, self.start_timer)
        else:
            self.end_quiz()

    def load_question(self):
        """Load the current question and its options."""
        question = self.questions[self.question_index]
        self.question_label.config(text=question["question"])
        for i, option in enumerate(question["options"]):
            self.option_buttons[i].config(text=option, state="normal")
        self.next_button.config(state="disabled")

    def check_answer(self, selected_index):
        """Check if the selected answer is correct."""
        correct_answer = self.questions[self.question_index]["correct"]
        selected_answer = self.option_buttons[selected_index].cget("text")
        if selected_answer == correct_answer:
            self.correct_answers += 1
            self.update_score()
        for button in self.option_buttons:
            button.config(state="disabled")
        self.next_button.config(state="normal")

    def update_score(self):
        """Update the score display."""
        self.score_label.config(text=f"Score: {self.correct_answers}/{self.total_marks}")

    def load_next_question(self):
        """Load the next question or end the quiz if no questions remain."""
        self.question_index += 1
        if self.question_index < len(self.questions):
            self.load_question()
        else:
            self.end_quiz()

    def end_quiz(self):
        """Display the quiz results and save to Excel."""
        incorrect_answers = self.total_marks - self.correct_answers
        messagebox.showinfo(
            "Quiz Over",
            f"Name: {self.name}\nSAP ID: {self.sapid}\n\n"
            f"Your final score is: {self.correct_answers}/{self.total_marks}\n"
            f"Correct Answers: {self.correct_answers}\n"
            f"Incorrect Answers: {incorrect_answers}"
        )
        self.save_to_excel()
        self.root.quit()

    def save_to_excel(self):
        """Save the quiz results to an Excel file."""
        file_name = r"./quiz.xlsx"

        # Check if the file already exists
        if os.path.exists(file_name):
            # Load the existing workbook
            workbook = openpyxl.load_workbook(file_name)
            sheet = workbook.active
        else:
            # Create a new workbook and add the header row
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Quiz Results"
            sheet.append(["Name", "SAP ID", "Total Marks", "Correct Answers", "Incorrect Answers"])

        # User data row
        incorrect_answers = self.total_marks - self.correct_answers
        sheet.append([self.name, self.sapid, self.total_marks, self.correct_answers, incorrect_answers])

        # Save to file
        workbook.save(file_name)
        messagebox.showinfo("Saved", f"Your results have been saved to '{file_name}'.")


if __name__ == "__main__":
    root = tk.Tk()
    app = QuizApp(root)
    root.mainloop()

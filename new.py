import os
import difflib
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
from tkinter.font import Font
from docx import Document
from pdfminer.high_level import extract_text
import re
from collections import Counter

# Simple sentence splitter (NO NLTK NEEDED!)
def split_into_sentences(text):
    """Split text into sentences without NLTK"""
    # Replace multiple sentence endings with single period
    text = re.sub(r'[.!?]+', '.', text)
    # Split by period and clean up
    sentences = [s.strip() for s in text.split('.') if s.strip()]
    return sentences

class DocumentComparator:
    def __init__(self, root):
        self.root = root
        self.root.title("Professional Document Comparator AI")
        self.root.geometry("1200x900")  # Increased height
        self.root.configure(bg='#f0f8ff')
        
        # Set up styles
        self.setup_styles()
        
        # Create UI elements
        self.create_widgets()
        
    def setup_styles(self):
        """Define custom styles for the application"""
        self.title_font = Font(family="Arial", size=18, weight="bold")
        self.label_font = Font(family="Arial", size=11)
        self.button_font = Font(family="Arial", size=11, weight="bold")
        self.result_font = Font(family="Arial", size=10)
        
    def create_widgets(self):
        """Create all UI elements"""
        # Title
        title_label = tk.Label(self.root, text="📄 Professional Document Comparator", 
                              font=self.title_font, bg='#f0f8ff', fg='#2c3e50')
        title_label.pack(pady=15)
        
        # Subtitle
        subtitle_label = tk.Label(self.root, text="Compare documents and find which one is more professional", 
                                 font=self.label_font, bg='#f0f8ff', fg='#7f8c8d')
        subtitle_label.pack(pady=5)
        
        # Main container
        main_frame = tk.Frame(self.root, bg='#f0f8ff')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # File selection section
        file_frame = tk.Frame(main_frame, bg='#f0f8ff')
        file_frame.pack(fill=tk.X, pady=15)
        
        # First file selection
        file1_frame = tk.Frame(file_frame, bg='#f0f8ff')
        file1_frame.pack(fill=tk.X, pady=8)
        
        tk.Label(file1_frame, text="📁 First Document:", font=self.label_font, 
                bg='#f0f8ff', width=18, anchor='w').pack(side=tk.LEFT)
        
        self.file1_path = tk.StringVar()
        file1_entry = tk.Entry(file1_frame, textvariable=self.file1_path, 
                               width=70, font=self.label_font)
        file1_entry.pack(side=tk.LEFT, padx=5)
        
        browse1_btn = tk.Button(file1_frame, text="Browse", command=self.browse_file1,
                               font=self.button_font, bg='#3498db', fg='white',
                               activebackground='#2980b9', relief=tk.FLAT, padx=20)
        browse1_btn.pack(side=tk.LEFT)
        
        # Second file selection
        file2_frame = tk.Frame(file_frame, bg='#f0f8ff')
        file2_frame.pack(fill=tk.X, pady=8)
        
        tk.Label(file2_frame, text="📁 Second Document:", font=self.label_font, 
                bg='#f0f8ff', width=18, anchor='w').pack(side=tk.LEFT)
        
        self.file2_path = tk.StringVar()
        file2_entry = tk.Entry(file2_frame, textvariable=self.file2_path, 
                               width=70, font=self.label_font)
        file2_entry.pack(side=tk.LEFT, padx=5)
        
        browse2_btn = tk.Button(file2_frame, text="Browse", command=self.browse_file2,
                               font=self.button_font, bg='#3498db', fg='white',
                               activebackground='#2980b9', relief=tk.FLAT, padx=20)
        browse2_btn.pack(side=tk.LEFT)
        
        # Button container
        button_frame = tk.Frame(main_frame, bg='#f0f8ff')
        button_frame.pack(fill=tk.X, pady=10)
        
        # Compare button
        compare_btn = tk.Button(button_frame, text="🔍 Compare Documents", 
                               command=self.compare_documents,
                               font=self.button_font, bg='#2ecc71', fg='white',
                               activebackground='#27ae60', relief=tk.FLAT, 
                               padx=20, pady=10)
        compare_btn.pack(side=tk.LEFT, padx=10)
        
        # Analyze professionalism button
        analyze_btn = tk.Button(button_frame, text="🎯 Which is More Professional?", 
                               command=self.analyze_professionalism,
                               font=self.button_font, bg='#e74c3c', fg='white',
                               activebackground='#c0392b', relief=tk.FLAT, 
                               padx=20, pady=10)
        analyze_btn.pack(side=tk.LEFT, padx=10)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", 
                                       length=100, mode="indeterminate")
        self.progress.pack(fill=tk.X, pady=5)
        
        # View mode selection
        view_frame = tk.Frame(main_frame, bg='#f0f8ff')
        view_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(view_frame, text="View Mode:", font=self.label_font, 
                bg='#f0f8ff').pack(side=tk.LEFT, padx=5)
        
        self.view_mode = tk.StringVar(value="simple")
        simple_radio = tk.Radiobutton(view_frame, text="Simple (Easy to Understand)", 
                                     variable=self.view_mode, value="simple",
                                     font=self.label_font, bg='#f0f8ff')
        simple_radio.pack(side=tk.LEFT, padx=10)
        
        technical_radio = tk.Radiobutton(view_frame, text="Technical (Detailed)", 
                                        variable=self.view_mode, value="technical",
                                        font=self.label_font, bg='#f0f8ff')
        technical_radio.pack(side=tk.LEFT, padx=10)
        
        # ====== FIXED: EXPORT BUTTONS AT TOP OF RESULTS ======
        # Create export toolbar ABOVE results
        export_toolbar = tk.Frame(main_frame, bg='#2c3e50', relief=tk.RAISED, bd=2)
        export_toolbar.pack(fill=tk.X, pady=5)
        
        # Export buttons in toolbar
        toolbar_label = tk.Label(export_toolbar, text="💾 EXPORT:", 
                                font=('Arial', 11, 'bold'), bg='#2c3e50', fg='white')
        toolbar_label.pack(side=tk.LEFT, padx=10, pady=5)
        
        # Main download button in toolbar
        download_btn = tk.Button(export_toolbar, text="📥 DOWNLOAD RESULTS", 
                                command=self.export_results,
                                font=('Arial', 11, 'bold'), bg='#e74c3c', fg='white',
                                activebackground='#c0392b', relief=tk.RAISED, bd=2,
                                padx=15, pady=5, cursor="hand2")
        download_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Clear button in toolbar
        clear_btn = tk.Button(export_toolbar, text="🗑️ CLEAR", 
                             command=self.clear_results,
                             font=('Arial', 11, 'bold'), bg='#95a5a6', fg='white',
                             activebackground='#7f8c8d', relief=tk.RAISED, bd=2,
                             padx=15, pady=5, cursor="hand2")
        clear_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Status in toolbar
        self.status_label = tk.Label(export_toolbar, text="✅ Ready to compare", 
                                    font=('Arial', 10, 'bold'), bg='#2c3e50', fg='#2ecc71')
        self.status_label.pack(side=tk.RIGHT, padx=10, pady=5)
        
        # Results section with side panel
        results_container = tk.Frame(main_frame, bg='#f0f8ff')
        results_container.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Left side - Results text
        left_frame = tk.Frame(results_container, bg='#f0f8ff')
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        tk.Label(left_frame, text="📊 Comparison Results:", 
                font=self.label_font, bg='#f0f8ff').pack(anchor='w')
        
        self.results_text = scrolledtext.ScrolledText(left_frame, wrap=tk.WORD, 
                                                     width=90, height=20, 
                                                     font=self.result_font)
        self.results_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Right side - Floating export panel
        right_panel = tk.Frame(results_container, bg='#34495e', relief=tk.RAISED, bd=3, width=200)
        right_panel.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        right_panel.pack_propagate(False)  # Keep fixed width
        
        # Panel header
        panel_header = tk.Label(right_panel, text="QUICK ACTIONS", 
                               font=('Arial', 12, 'bold'), bg='#34495e', fg='white')
        panel_header.pack(pady=10)
        
        # Separator
        separator1 = tk.Frame(right_panel, height=2, bg='#2c3e50')
        separator1.pack(fill=tk.X, padx=10, pady=5)
        
        # Export button in side panel
        side_export_btn = tk.Button(right_panel, text="📥\nDOWNLOAD\nRESULTS", 
                                   command=self.export_results,
                                   font=('Arial', 11, 'bold'), bg='#e74c3c', fg='white',
                                   activebackground='#c0392b', relief=tk.RAISED, bd=2,
                                   width=15, height=3, cursor="hand2")
        side_export_btn.pack(pady=10, padx=10)
        
        # Separator
        separator2 = tk.Frame(right_panel, height=2, bg='#2c3e50')
        separator2.pack(fill=tk.X, padx=10, pady=5)
        
        # Clear button in side panel
        side_clear_btn = tk.Button(right_panel, text="🗑️\nCLEAR\nRESULTS", 
                                  command=self.clear_results,
                                  font=('Arial', 11, 'bold'), bg='#95a5a6', fg='white',
                                  activebackground='#7f8c8d', relief=tk.RAISED, bd=2,
                                  width=15, height=3, cursor="hand2")
        side_clear_btn.pack(pady=10, padx=10)
        
        # Panel footer
        panel_footer = tk.Label(right_panel, text="Click to save\nyour comparison\nreport", 
                               font=('Arial', 9), bg='#34495e', fg='#ecf0f1')
        panel_footer.pack(pady=10)
        
        # Configure text tags for highlighting
        self.results_text.tag_config("added", background="#d4edda", font=('Arial', 10, 'bold'))
        self.results_text.tag_config("removed", background="#f8d7da", font=('Arial', 10, 'bold'))
        self.results_text.tag_config("unchanged", foreground="#2c3e50")
        self.results_text.tag_config("header", font=('Arial', 11, 'bold'), foreground='#2c3e50')
        self.results_text.tag_config("summary", font=('Arial', 10, 'bold'), background='#e8f4fd')
        self.results_text.tag_config("professional", font=('Arial', 10, 'bold'), foreground='#8e44ad')
        
    def browse_file1(self):
        """Open file dialog to select the first document"""
        file_path = filedialog.askopenfilename(
            title="Select First Document",
            filetypes=[
                ("All Supported Files", "*.pdf;*.docx;*.doc"),
                ("PDF Files", "*.pdf"), 
                ("Word Files", "*.docx;*.doc"),
                ("All Files", "*.*")
            ]
        )
        if file_path:
            self.file1_path.set(file_path)
            self.status_label.config(text="✅ File 1 selected", fg='#2ecc71')
    
    def browse_file2(self):
        """Open file dialog to select the second document"""
        file_path = filedialog.askopenfilename(
            title="Select Second Document",
            filetypes=[
                ("All Supported Files", "*.pdf;*.docx;*.doc"),
                ("PDF Files", "*.pdf"), 
                ("Word Files", "*.docx;*.doc"),
                ("All Files", "*.*")
            ]
        )
        if file_path:
            self.file2_path.set(file_path)
            self.status_label.config(text="✅ File 2 selected", fg='#2ecc71')
    
    def extract_text_from_file(self, file_path):
        """Extract text from PDF or Word document with better cleaning"""
        try:
            file_extension = os.path.splitext(file_path)[1].lower()
            
            if file_extension == '.pdf':
                text = extract_text(file_path)
                # Better text cleaning for PDF
                text = re.sub(r'\s+', ' ', text)  # Replace multiple spaces
                text = re.sub(r'\n+', ' ', text)  # Replace newlines
                text = text.strip()
                return text
            elif file_extension in ['.docx', '.doc']:
                doc = Document(file_path)
                text = '\n'.join([para.text for para in doc.paragraphs])
                # Better text cleaning for Word
                text = re.sub(r'\s+', ' ', text)
                text = text.strip()
                return text
            else:
                return "Unsupported file format. Please use PDF or Word (.docx, .doc) files."
        except Exception as e:
            return f"Error reading file: {str(e)}"
    
    def compare_documents(self):
        """Compare the two selected documents"""
        file1 = self.file1_path.get()
        file2 = self.file2_path.get()
        
        if not file1 or not file2:
            messagebox.showwarning("Warning", "Please select both documents to compare.")
            return
        
        # Check if same file
        if os.path.abspath(file1) == os.path.abspath(file2):
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, "="*60 + "\n", "header")
            self.results_text.insert(tk.END, "⚠️  SAME FILE DETECTED!\n", "header")
            self.results_text.insert(tk.END, "="*60 + "\n\n", "header")
            self.results_text.insert(tk.END, "You have selected the same file for both documents.\n", "summary")
            self.results_text.insert(tk.END, "Please select two different files to compare.\n", "summary")
            self.status_label.config(text="⚠️ Same file selected!", fg='#e74c3c')
            return
        
        # Start progress bar
        self.progress.start()
        self.status_label.config(text="🔄 Comparing...", fg='#f39c12')
        self.root.update()
        
        # Extract text from both files
        text1 = self.extract_text_from_file(file1)
        text2 = self.extract_text_from_file(file2)
        
        # Clear previous results
        self.results_text.delete(1.0, tk.END)
        
        # Check if view mode is simple or technical
        if self.view_mode.get() == "simple":
            self.show_simple_comparison(text1, text2, file1, file2)
        else:
            self.show_technical_comparison(text1, text2, file1, file2)
        
        # Stop progress bar
        self.progress.stop()
        self.status_label.config(text="✅ Done! Click DOWNLOAD to save", fg='#2ecc71')
    
    def analyze_professionalism(self):
        """Analyze which document is more professional"""
        file1 = self.file1_path.get()
        file2 = self.file2_path.get()
        
        if not file1 or not file2:
            messagebox.showwarning("Warning", "Please select both documents to analyze.")
            return
        
        # Check if same file
        if os.path.abspath(file1) == os.path.abspath(file2):
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, "="*60 + "\n", "header")
            self.results_text.insert(tk.END, "⚠️  SAME FILE DETECTED!\n", "header")
            self.results_text.insert(tk.END, "="*60 + "\n\n", "header")
            self.results_text.insert(tk.END, "Please select two different files to analyze.\n", "summary")
            self.status_label.config(text="⚠️ Same file selected!", fg='#e74c3c')
            return
        
        # Start progress bar
        self.progress.start()
        self.status_label.config(text="🔄 Analyzing...", fg='#f39c12')
        self.root.update()
        
        # Extract text from both files
        text1 = self.extract_text_from_file(file1)
        text2 = self.extract_text_from_file(file2)
        
        # Clear previous results
        self.results_text.delete(1.0, tk.END)
        
        # Add header
        self.results_text.insert(tk.END, "="*60 + "\n", "header")
        self.results_text.insert(tk.END, f"PROFESSIONALISM ANALYSIS REPORT\n", "header")
        self.results_text.insert(tk.END, "="*60 + "\n\n", "header")
        
        # Analyze both documents
        score1 = self.calculate_professional_score(text1, os.path.basename(file1))
        score2 = self.calculate_professional_score(text2, os.path.basename(file2))
        
        # Show comparison
        self.results_text.insert(tk.END, "📊 PROFESSIONALISM SCORES:\n", "summary")
        self.results_text.insert(tk.END, "-"*50 + "\n", "summary")
        self.results_text.insert(tk.END, f"📄 {os.path.basename(file1)}: {score1}/100\n", "summary")
        self.results_text.insert(tk.END, f"📄 {os.path.basename(file2)}: {score2}/100\n\n", "summary")
        
        # Determine winner
        if score1 > score2:
            winner = os.path.basename(file1)
            winner_score = score1
            loser_score = score2
        elif score2 > score1:
            winner = os.path.basename(file2)
            winner_score = score2
            loser_score = score1
        else:
            winner = "Both documents"
            winner_score = score1
            loser_score = score2
        
        # Show conclusion
        self.results_text.insert(tk.END, "🏆 CONCLUSION:\n", "professional")
        self.results_text.insert(tk.END, "-"*50 + "\n", "professional")
        if score1 == score2:
            self.results_text.insert(tk.END, "Both documents have EQUAL PROFESSIONALISM!\n\n", "professional")
            self.results_text.insert(tk.END, f"Both scored {winner_score}/100.\n\n", "professional")
        else:
            self.results_text.insert(tk.END, f"'{winner}' is MORE PROFESSIONAL!\n\n", "professional")
            self.results_text.insert(tk.END, f"It scored {winner_score}/100 while the other document scored {loser_score}/100.\n\n", "professional")
        
        # Show improvement suggestions
        self.results_text.insert(tk.END, "💡 IMPROVEMENT SUGGESTIONS:\n", "header")
        self.results_text.insert(tk.END, "-"*50 + "\n", "header")
        
        if score1 < 80:
            self.results_text.insert(tk.END, f"For '{os.path.basename(file1)}':\n", "added")
            self.results_text.insert(tk.END, "• Use more professional language\n", "added")
            self.results_text.insert(tk.END, "• Check for grammar and spelling errors\n", "added")
            self.results_text.insert(tk.END, "• Improve document structure\n", "added")
            self.results_text.insert(tk.END, "• Add more detailed information\n\n", "added")
        
        if score2 < 80:
            self.results_text.insert(tk.END, f"For '{os.path.basename(file2)}':\n", "added")
            self.results_text.insert(tk.END, "• Use more professional language\n", "added")
            self.results_text.insert(tk.END, "• Check for grammar and spelling errors\n", "added")
            self.results_text.insert(tk.END, "• Improve document structure\n", "added")
            self.results_text.insert(tk.END, "• Add more detailed information\n\n", "added")
        
        # Stop progress bar
        self.progress.stop()
        self.status_label.config(text="✅ Done! Click DOWNLOAD to save", fg='#2ecc71')
    
    def calculate_professional_score(self, text, filename):
        """Calculate a professionalism score for the document"""
        score = 50  # Base score
        
        # Check for professional sections (for CVs)
        professional_sections = [
            'experience', 'education', 'skills', 'projects', 'certifications',
            'achievements', 'objective', 'summary', 'references', 'contact'
        ]
        
        text_lower = text.lower()
        section_count = 0
        
        for section in professional_sections:
            if section in text_lower:
                section_count += 1
                score += 3
        
        # Check document length (longer is usually more detailed)
        word_count = len(text.split())
        if word_count > 200:
            score += 5
        elif word_count > 100:
            score += 2
        
        # Check for professional keywords
        professional_keywords = [
            'managed', 'developed', 'implemented', 'achieved', 'led',
            'coordinated', 'designed', 'created', 'improved', 'increased',
            'reduced', 'optimized', 'analyzed', 'researched', 'presented'
        ]
        
        keyword_count = 0
        for keyword in professional_keywords:
            if keyword in text_lower:
                keyword_count += 1
        
        score += min(keyword_count * 2, 15)  # Max 15 points for keywords
        
        # Check for unprofessional elements
        unprofessional_indicators = [
            'lol', 'omg', 'btw', 'idk', 'tbh', 'fyi', 'asap',
            'cool', 'awesome', 'amazing', 'stuff', 'things'
        ]
        
        for indicator in unprofessional_indicators:
            if indicator in text_lower:
                score -= 2
        
        # Check for proper sentence structure using our simple splitter
        sentences = split_into_sentences(text)
        if sentences:
            avg_sentence_length = sum(len(s.split()) for s in sentences) / len(sentences)
            if 10 <= avg_sentence_length <= 25:
                score += 5  # Good sentence length
        
        # Check for contact information
        if re.search(r'[\w\.-]+@[\w\.-]+\.\w+', text):  # Email
            score += 5
        if re.search(r'(\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', text):  # Phone
            score += 5
        
        # Ensure score is within 0-100 range
        return max(0, min(100, score))
    
    def show_simple_comparison(self, text1, text2, file1, file2):
        """Show comparison in simple, easy-to-understand format"""
        # Add header
        self.results_text.insert(tk.END, "="*60 + "\n", "header")
        self.results_text.insert(tk.END, f"DOCUMENT COMPARISON REPORT\n", "header")
        self.results_text.insert(tk.END, "="*60 + "\n\n", "header")
        
        self.results_text.insert(tk.END, f"📄 Document 1: {os.path.basename(file1)}\n", "header")
        self.results_text.insert(tk.END, f"📄 Document 2: {os.path.basename(file2)}\n\n", "header")
        
        # Split text into sentences using our simple splitter
        sentences1 = split_into_sentences(text1)
        sentences2 = split_into_sentences(text2)
        
        # Find similarities and differences
        common_sentences = []
        only_in_doc1 = []
        only_in_doc2 = []
        
        # Use SequenceMatcher to find similar sentences
        matcher = difflib.SequenceMatcher(None, sentences1, sentences2)
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                for i in range(i1, i2):
                    common_sentences.append(sentences1[i])
            elif tag == 'delete':
                for i in range(i1, i2):
                    only_in_doc1.append(sentences1[i])
            elif tag == 'insert':
                for j in range(j1, j2):
                    only_in_doc2.append(sentences2[j])
            elif tag == 'replace':
                for i in range(i1, i2):
                    only_in_doc1.append(sentences1[i])
                for j in range(j1, j2):
                    only_in_doc2.append(sentences2[j])
        
        # Show summary
        self.results_text.insert(tk.END, "📋 SUMMARY:\n", "summary")
        self.results_text.insert(tk.END, f"• Common content: {len(common_sentences)} sentences\n", "summary")
        self.results_text.insert(tk.END, f"• Only in Document 1: {len(only_in_doc1)} sentences\n", "summary")
        self.results_text.insert(tk.END, f"• Only in Document 2: {len(only_in_doc2)} sentences\n\n", "summary")
        
        # Show common content
        if common_sentences:
            self.results_text.insert(tk.END, "✅ CONTENT FOUND IN BOTH DOCUMENTS:\n", "header")
            self.results_text.insert(tk.END, "-"*50 + "\n", "header")
            for sentence in common_sentences[:10]:  # Show first 10 common sentences
                self.results_text.insert(tk.END, f"• {sentence}.\n", "unchanged")
            if len(common_sentences) > 10:
                self.results_text.insert(tk.END, f"... and {len(common_sentences) - 10} more common sentences\n", "unchanged")
            self.results_text.insert(tk.END, "\n")
        
        # Show content only in document 1
        if only_in_doc1:
            self.results_text.insert(tk.END, "❌ CONTENT ONLY IN DOCUMENT 1 (Removed):\n", "removed")
            self.results_text.insert(tk.END, "-"*50 + "\n", "removed")
            for sentence in only_in_doc1[:10]:  # Show first 10 unique sentences
                self.results_text.insert(tk.END, f"• {sentence}.\n", "removed")
            if len(only_in_doc1) > 10:
                self.results_text.insert(tk.END, f"... and {len(only_in_doc1) - 10} more sentences\n", "removed")
            self.results_text.insert(tk.END, "\n")
        
        # Show content only in document 2
        if only_in_doc2:
            self.results_text.insert(tk.END, "✅ CONTENT ONLY IN DOCUMENT 2 (Added):\n", "added")
            self.results_text.insert(tk.END, "-"*50 + "\n", "added")
            for sentence in only_in_doc2[:10]:  # Show first 10 unique sentences
                self.results_text.insert(tk.END, f"• {sentence}.\n", "added")
            if len(only_in_doc2) > 10:
                self.results_text.insert(tk.END, f"... and {len(only_in_doc2) - 10} more sentences\n", "added")
        
        # Add conclusion
        self.results_text.insert(tk.END, "\n" + "="*60 + "\n", "header")
        similarity = (len(common_sentences) / max(len(sentences1), len(sentences2))) * 100 if sentences1 or sentences2 else 0
        self.results_text.insert(tk.END, f"CONCLUSION: Documents are {similarity:.1f}% similar\n", "summary")
        self.results_text.insert(tk.END, "="*60 + "\n", "header")
    
    def show_technical_comparison(self, text1, text2, file1, file2):
        """Show technical diff comparison"""
        # Split text into lines for comparison
        lines1 = text1.splitlines()
        lines2 = text2.splitlines()
        
        # Add header
        self.results_text.insert(tk.END, f"Comparing: {os.path.basename(file1)} vs {os.path.basename(file2)}\n\n", "header")
        
        # Generate a unified diff
        unified_diff = list(difflib.unified_diff(
            lines1, lines2, 
            fromfile=os.path.basename(file1), 
            tofile=os.path.basename(file2),
            lineterm='', n=3
        ))
        
        # Process and display the diff
        for line in unified_diff:
            if line.startswith('---') or line.startswith('+++'):
                self.results_text.insert(tk.END, line + '\n', "header")
            elif line.startswith('@@'):
                self.results_text.insert(tk.END, line + '\n', "changed")
            elif line.startswith('-'):
                self.results_text.insert(tk.END, line + '\n', "removed")
            elif line.startswith('+'):
                self.results_text.insert(tk.END, line + '\n', "added")
            else:
                self.results_text.insert(tk.END, line + '\n')
        
        # Show summary
        added_count = sum(1 for line in unified_diff if line.startswith('+'))
        removed_count = sum(1 for line in unified_diff if line.startswith('-'))
        
        self.results_text.insert(tk.END, f"\n\nSummary: {added_count} additions, {removed_count} removals\n", "summary")
    
    def export_results(self):
        """Export comparison results to a text file"""
        content = self.results_text.get(1.0, tk.END)
        if not content.strip():
            messagebox.showwarning("Warning", "No results to export. Please compare documents first.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[
                ("Text Files", "*.txt"),
                ("All Files", "*.*")
            ],
            title="Save Comparison Results"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                messagebox.showinfo("Success", f"✅ Results saved successfully!\n\nFile location: {file_path}")
                self.status_label.config(text="✅ Saved successfully!", fg='#2ecc71')
            except Exception as e:
                messagebox.showerror("Error", f"❌ Failed to save results: {str(e)}")
                self.status_label.config(text="❌ Save failed!", fg='#e74c3c')
    
    def clear_results(self):
        """Clear the results text area"""
        self.results_text.delete(1.0, tk.END)
        self.status_label.config(text="✅ Results cleared", fg='#95a5a6')

if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentComparator(root)
    root.mainloop()
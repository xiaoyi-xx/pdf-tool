import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
from pathlib import Path

# 导入合并工具模块
try:
    from pdf_merge_tool import PDFMergeTool
    from pdf_split_tool import PDFSplitTool
    from pdf_compress_tool import PDFCompressTool
    from pdf_encrypt_decrypt_tool import PDFEncryptDecryptTool
    from pdf_to_word_tool import PDFToWordTool
    from pdf_image_converter_tool import PDFImageConverterTool
    from pdf_to_text_tool import PDFToTextTool
    from pdf_rotate_tool import PDFRotateTool
    from pdf_metadata_tool import PDFMetadataTool
    from pdf_watermark_tool import PDFWatermarkTool
    from pdf_header_footer_tool import PDFHeaderFooterTool
    from word_to_pdf_tool import WordToPDFTool
    from excel_to_pdf_tool import ExcelToPDFTool
    from ppt_to_pdf_tool import PPTToPDFTool
    from pdf_bookmark_tool import PDFBookmarkTool
    from pdf_annotation_tool import PDFAnnotationTool
    from pdf_form_tool import PDFFormTool
    from pdf_batch_tool import PDFBatchTool
    from pdf_compare_tool import PDFCompareTool
    from pdf_ocr_tool import PDFOCRTool
    from pdf_optimize_tool import PDFOptimizeTool
    from pdf_signature_tool import PDFSignatureTool
except ImportError:
    # 如果导入失败，创建一个占位符类
    class PDFMergeTool:
        def __init__(self, parent_frame, file_list=None):
            self.parent = parent_frame
            self.create_placeholder_interface("PDF合并工具", "将多个PDF文件合并为一个PDF文件")
            self.create_placeholder_interface("PDF分割工具", "将多个PDF文件分割为多个PDF文件")
            self.create_placeholder_interface("PDF压缩工具", "使用多种算法减小PDF文件大小，优化存储和传输")
            self.create_placeholder_interface("PDF加密/解密工具", "为PDF文件添加密码保护或移除密码保护")
            self.create_placeholder_interface("PDF转Word工具", "将PDF文件转换为可编辑的Word文档")
            self.create_placeholder_interface("PDF与图片互转工具", "PDF转图片或将多张图片合并为PDF")
            self.create_placeholder_interface("PDF转文本工具", "从PDF文件中提取文本内容")
            self.create_placeholder_interface("PDF旋转工具", "旋转PDF页面方向")
        
        def create_placeholder_interface(self, title, description):
            for widget in self.parent.winfo_children():
                widget.destroy()
            
            title_label = ttk.Label(self.parent, text=title, font=('Arial', 16, 'bold'))
            title_label.pack(pady=(10, 5))
            
            desc_label = ttk.Label(self.parent, text=description, font=('Arial', 12))
            desc_label.pack(pady=(0, 20))
            
            content_frame = ttk.LabelFrame(self.parent, text="功能区域")
            content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            placeholder = ttk.Label(content_frame, text="此功能正在开发中...\n\n请等待后续更新", 
                                   justify=tk.CENTER, font=('Arial', 12))
            placeholder.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)
    
    # 为其他未导入的工具创建占位符类
    class PDFToTextTool(PDFMergeTool):
        pass
    
    class PDFRotateTool(PDFMergeTool):
        pass
    
    class PDFMetadataTool(PDFMergeTool):
        pass
    
    class PDFWatermarkTool(PDFMergeTool):
        pass
    
    class PDFHeaderFooterTool(PDFMergeTool):
        pass
    
    class WordToPDFTool(PDFMergeTool):
        pass
    
    class ExcelToPDFTool(PDFMergeTool):
        pass
    
    class PPTToPDFTool(PDFMergeTool):
        pass
    
    class PDFBookmarkTool(PDFMergeTool):
        pass
    
    class PDFAnnotationTool(PDFMergeTool):
        pass
    
    class PDFFormTool(PDFMergeTool):
        pass
    
    class PDFBatchTool(PDFMergeTool):
        pass
    
    class PDFCompareTool(PDFMergeTool):
        pass
    
    class PDFOCRTool(PDFMergeTool):
        pass
    
    class PDFOptimizeTool(PDFMergeTool):
        pass
    
    class PDFSignatureTool(PDFMergeTool):
        pass

class PDFToolbox:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF工具箱 - 多功能PDF处理工具")
        self.root.geometry("1000x700")
        self.root.resizable(True, True)
        
        # 确保主窗口占满整个屏幕空间
        self.root.pack_propagate(False)
        
        # 存储选中的文件
        self.current_files = []
        
        # 设置样式
        self.setup_styles()
        
        # 创建主界面
        self.create_main_interface()
        
        # 初始化功能模块
        self.function_frames = {}
        self.setup_function_frames()
        
    def setup_styles(self):
        """设置界面样式"""
        style = ttk.Style()
        style.configure('Title.TLabel', font=('Arial', 18, 'bold'))
        style.configure('Subtitle.TLabel', font=('Arial', 12, 'bold'))
        style.configure('Action.TButton', font=('Arial', 10))
        style.configure('Function.TButton', font=('Arial', 9), width=15)
        
    def create_main_interface(self):
        """创建主界面"""
        # 主容器
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main_container.pack_propagate(False)  # 防止子组件改变容器大小
        
        # 标题
        title_label = ttk.Label(main_container, text="PDF工具箱", style='Title.TLabel')
        title_label.pack(pady=(0, 20))
        
        # 创建左右分栏
        paned_window = ttk.PanedWindow(main_container, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True, pady=(0, 0), padx=(0, 0))
        
        # 左侧功能选择区域 - 使用Frame包装以便添加滚动条
        left_container = ttk.Frame(paned_window, width=220)
        left_container.pack_propagate(False)  # 防止子组件改变容器大小
        paned_window.add(left_container, weight=0)
        
        # 右侧功能区域
        self.right_frame = ttk.Frame(paned_window)
        self.right_frame.pack_propagate(False)  # 防止子组件改变容器大小
        paned_window.add(self.right_frame, weight=1)
        
        # 在左侧容器中创建文件列表区域
        self.create_file_list_section(left_container)
        
        # 在左侧容器中创建可滚动的功能菜单
        self.create_scrollable_function_menu(left_container)
        
        # 在右侧显示欢迎界面
        self.show_welcome_screen()
        
    def create_file_list_section(self, parent):
        """创建文件列表区域"""
        # 文件列表容器
        file_container = ttk.LabelFrame(parent, text="文件操作")
        file_container.pack(fill=tk.BOTH, expand=False, padx=5, pady=5)
        
        # 文件操作按钮
        file_buttons_frame = ttk.Frame(file_container)
        file_buttons_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 添加文件按钮
        add_file_btn = ttk.Button(file_buttons_frame, text="添加PDF文件", 
                                 command=self.add_pdf_files)
        add_file_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 移除文件按钮
        remove_file_btn = ttk.Button(file_buttons_frame, text="移除选中", 
                                    command=self.remove_selected_files)
        remove_file_btn.pack(side=tk.LEFT, padx=5)
        
        # 清空列表按钮
        clear_file_btn = ttk.Button(file_buttons_frame, text="清空列表", 
                                   command=self.clear_files)
        clear_file_btn.pack(side=tk.LEFT, padx=5)
        
        # 文件列表框
        list_frame = ttk.Frame(file_container)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建文件列表框和滚动条
        self.file_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, height=6)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        file_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, 
                                      command=self.file_listbox.yview)
        file_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.configure(yscrollcommand=file_scrollbar.set)
    
    def create_scrollable_function_menu(self, parent):
        """创建可滚动的功能菜单"""
        # 创建一个Frame来包含Canvas和滚动条
        menu_container = ttk.Frame(parent)
        menu_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建Canvas和滚动条
        canvas = tk.Canvas(menu_container, width=220)
        scrollbar = ttk.Scrollbar(menu_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        # 配置Canvas
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 布局Canvas和滚动条
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 构建功能菜单
        self.create_function_menu(scrollable_frame)
        
        # 绑定鼠标滚轮事件
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind("<MouseWheel>", _on_mousewheel)
        scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
        
    def create_function_menu(self, parent):
        """创建功能菜单"""
        
        
        # 功能分类区域
        function_frame = ttk.LabelFrame(parent, text="PDF功能")
        function_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 基础操作
        basic_frame = ttk.LabelFrame(function_frame, text="基础操作")
        basic_frame.pack(fill=tk.X, padx=5, pady=5)
        
        functions_basic = [
            ("PDF合并", self.show_merge_tool),
            ("PDF分割", self.show_split_tool),
            ("PDF压缩", self.show_compress_tool),
            ("PDF加密/解密", self.show_encrypt_tool),
            # ("PDF解密", self.show_decrypt_tool),
        ]
        
        for text, command in functions_basic:
            btn = ttk.Button(basic_frame, text=text, command=command, style='Function.TButton')
            btn.pack(fill=tk.X, padx=5, pady=2)
        
        # 转换功能
        convert_frame = ttk.LabelFrame(function_frame, text="转换功能")
        convert_frame.pack(fill=tk.X, padx=5, pady=5)
        
        functions_convert = [
            ("PDF转Word", self.show_pdf_to_word_tool),
            ("PDF&图片互转", self.show_pdf_to_image_tool),
            ("PDF转文本", self.show_pdf_to_text_tool),
            ("Word转PDF", self.show_word_to_pdf_tool),
            ("Excel转PDF", self.show_excel_to_pdf_tool),
            ("PPT转PDF", self.show_ppt_to_pdf_tool),
        ]
        
        for text, command in functions_convert:
            btn = ttk.Button(convert_frame, text=text, command=command, style='Function.TButton')
            btn.pack(fill=tk.X, padx=5, pady=2)
        
        # 编辑功能
        edit_frame = ttk.LabelFrame(function_frame, text="编辑功能")
        edit_frame.pack(fill=tk.X, padx=5, pady=5)
        
        functions_edit = [
            ("PDF水印", self.show_watermark_tool),
            ("PDF旋转", self.show_rotate_tool),
            ("PDF页眉页脚", self.show_header_footer_tool),
            ("PDF书签", self.show_bookmark_tool),
            ("PDF注释", self.show_annotation_tool),
            ("PDF表单", self.show_form_tool),
        ]
        
        for text, command in functions_edit:
            btn = ttk.Button(edit_frame, text=text, command=command, style='Function.TButton')
            btn.pack(fill=tk.X, padx=5, pady=2)
        
        # 高级功能
        advanced_frame = ttk.LabelFrame(function_frame, text="高级功能")
        advanced_frame.pack(fill=tk.X, padx=5, pady=5)
        
        functions_advanced = [
            ("PDF比较", self.show_compare_tool),
            ("OCR识别", self.show_ocr_tool),
            ("批量处理", self.show_batch_tool),
            ("PDF优化", self.show_optimize_tool),
            ("PDF签名", self.show_signature_tool),
            ("元数据编辑", self.show_metadata_tool),
        ]
        
        for text, command in functions_advanced:
            btn = ttk.Button(advanced_frame, text=text, command=command, style='Function.TButton')
            btn.pack(fill=tk.X, padx=5, pady=2)
    
    def setup_function_frames(self):
        """初始化所有功能模块的框架"""
        # 这里将初始化所有功能模块的框架
        # 目前先创建空框架，后续每个功能模块会替换这些框架的内容
        
        # 基础操作
        self.function_frames["merge"] = self.create_function_frame("PDF合并工具", "将多个PDF文件合并为一个PDF文件")
        self.function_frames["split"] = self.create_function_frame("PDF分割工具", "将PDF文件分割为多个部分")
        self.function_frames["compress"] = self.create_function_frame("PDF压缩工具", "减小PDF文件大小")
        self.function_frames["encrypt"] = self.create_function_frame("PDF加密工具", "为PDF文件添加密码保护")
        self.function_frames["decrypt"] = self.create_function_frame("PDF解密工具", "移除PDF文件的密码保护")
        
        # 转换功能
        self.function_frames["pdf_to_word"] = self.create_function_frame("PDF转Word工具", "将PDF文件转换为可编辑的Word文档")
        self.function_frames["pdf_to_image"] = self.create_function_frame("PDF转图片工具", "将PDF页面转换为图片格式")
        self.function_frames["image_to_pdf"] = self.create_function_frame("图片转PDF工具", "将多张图片合并为PDF文件")
        self.function_frames["pdf_to_text"] = self.create_function_frame("PDF转文本工具", "从PDF文件中提取文本内容")
        self.function_frames["word_to_pdf"] = self.create_function_frame("Word转PDF工具", "将Word文档转换为PDF格式")
        self.function_frames["excel_to_pdf"] = self.create_function_frame("Excel转PDF工具", "将Excel表格转换为PDF格式")
        self.function_frames["ppt_to_pdf"] = self.create_function_frame("PPT转PDF工具", "将PowerPoint演示文稿转换为PDF格式")
        
        # 编辑功能
        self.function_frames["watermark"] = self.create_function_frame("PDF水印工具", "为PDF文件添加文字或图片水印")
        self.function_frames["rotate"] = self.create_function_frame("PDF旋转工具", "旋转PDF页面方向")
        self.function_frames["header_footer"] = self.create_function_frame("PDF页眉页脚工具", "为PDF文件添加页眉和页脚")
        self.function_frames["bookmark"] = self.create_function_frame("PDF书签工具", "为PDF文件添加或编辑书签")
        self.function_frames["annotation"] = self.create_function_frame("PDF注释工具", "为PDF文件添加注释、高亮和下划线")
        self.function_frames["form"] = self.create_function_frame("PDF表单工具", "填写或创建PDF表单")
        
        # 高级功能
        self.function_frames["compare"] = self.create_function_frame("PDF比较工具", "比较两个PDF文件的差异")
        self.function_frames["ocr"] = self.create_function_frame("OCR识别工具", "识别扫描PDF中的文字内容")
        self.function_frames["batch"] = self.create_function_frame("批量处理工具", "批量处理多个PDF文件")
        self.function_frames["optimize"] = self.create_function_frame("PDF优化工具", "优化PDF文件结构和性能")
        self.function_frames["signature"] = self.create_function_frame("PDF签名工具", "为PDF文件添加数字签名")
        self.function_frames["metadata"] = self.create_function_frame("元数据编辑工具", "编辑PDF文件的元数据信息")
    
    def create_function_frame(self, title, description):
        """创建功能框架的通用方法"""
        frame = ttk.Frame(self.right_frame)
        
        # 标题
        title_label = ttk.Label(frame, text=title, style='Title.TLabel')
        title_label.pack(pady=(10, 5))
        
        # 描述
        desc_label = ttk.Label(frame, text=description, style='Subtitle.TLabel')
        desc_label.pack(pady=(0, 20))
        
        # 功能区域
        content_frame = ttk.LabelFrame(frame, text="功能区域")
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 占位文本 - 实际功能将在这里实现
        placeholder = ttk.Label(content_frame, text="此功能正在开发中...\n\n请等待后续更新", 
                               justify=tk.CENTER, font=('Arial', 12))
        placeholder.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)
        
        # 操作按钮区域
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 执行按钮
        execute_btn = ttk.Button(button_frame, text="执行操作", 
                                command=lambda: self.function_not_implemented(title),
                                style='Action.TButton')
        execute_btn.pack(side=tk.RIGHT, padx=5)
        
        # 重置按钮
        reset_btn = ttk.Button(button_frame, text="重置", 
                              command=lambda: self.function_not_implemented(title),
                              style='Action.TButton')
        reset_btn.pack(side=tk.RIGHT, padx=5)
        
        return frame
    
    def show_welcome_screen(self):
        """显示欢迎界面"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建欢迎界面
        welcome_frame = ttk.Frame(self.right_frame)
        welcome_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 欢迎标题
        welcome_title = ttk.Label(welcome_frame, text="欢迎使用PDF工具箱", style='Title.TLabel')
        welcome_title.pack(pady=(50, 20))
        
        # 欢迎描述
        welcome_desc = ttk.Label(welcome_frame, 
                                text="这是一个多功能的PDF处理工具集，提供各种PDF文件的处理功能。\n\n"
                                     "请从左侧菜单选择您需要的功能，然后添加PDF文件开始使用。",
                                justify=tk.CENTER, font=('Arial', 12))
        welcome_desc.pack(pady=(0, 30))
        
        # 功能列表
        features_frame = ttk.LabelFrame(welcome_frame, text="可用功能")
        features_frame.pack(fill=tk.BOTH, expand=True, padx=50, pady=10)
        
        # 功能分类显示
        features_text = """
        • 基础操作: PDF合并、分割、压缩、加密、解密
        • 转换功能: PDF转Word、PDF转图片、图片转PDF、PDF转文本、Office转PDF
        • 编辑功能: PDF水印、PDF旋转、页眉页脚、书签管理、注释、表单
        • 高级功能: PDF比较、OCR识别、批量处理、PDF优化、数字签名、元数据编辑
        """
        
        # 使用Text组件替代Label，支持滚动
        features_text_widget = tk.Text(features_frame, wrap=tk.WORD, font=('Arial', 11), height=8)
        features_text_widget.insert(tk.END, features_text)
        features_text_widget.config(state=tk.DISABLED)  # 设置为只读
        
        # 添加滚动条
        features_scrollbar = ttk.Scrollbar(features_frame, orient=tk.VERTICAL, command=features_text_widget.yview)
        features_text_widget.configure(yscrollcommand=features_scrollbar.set)
        
        features_text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=20, pady=20)
        features_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=20)
        
        # 使用提示
        tip_frame = ttk.LabelFrame(welcome_frame, text="使用提示")
        tip_frame.pack(fill=tk.X, padx=50, pady=10)
        
        tips_text = """
        1. 首先从左侧"文件操作"区域添加PDF文件
        2. 然后从左侧功能菜单中选择所需功能
        3. 根据功能要求设置相关参数
        4. 点击"执行操作"按钮完成处理
        """
        
        # 使用Text组件替代Label，支持滚动
        tips_text_widget = tk.Text(tip_frame, wrap=tk.WORD, font=('Arial', 10), height=4)
        tips_text_widget.insert(tk.END, tips_text)
        tips_text_widget.config(state=tk.DISABLED)  # 设置为只读
        
        # 添加滚动条
        tips_scrollbar = ttk.Scrollbar(tip_frame, orient=tk.VERTICAL, command=tips_text_widget.yview)
        tips_text_widget.configure(yscrollcommand=tips_scrollbar.set)
        
        tips_text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=20, pady=10)
        tips_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=10)
    
    def add_pdf_files(self):
        """添加PDF文件"""
        files = filedialog.askopenfilenames(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        for file in files:
            if file not in self.current_files:
                self.current_files.append(file)
                self.file_listbox.insert(tk.END, os.path.basename(file))
        
        if files:
            messagebox.showinfo("成功", f"已添加 {len(files)} 个PDF文件")
    
    def remove_selected_files(self):
        """从文件列表中移除选中的文件"""
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "请先选择要移除的文件")
            return
            
        for index in reversed(selected_indices):
            self.file_listbox.delete(index)
            self.current_files.pop(index)
        
        messagebox.showinfo("成功", f"已移除 {len(selected_indices)} 个文件")
    
    def clear_files(self):
        """清除文件列表"""
        if not self.current_files:
            messagebox.showinfo("提示", "文件列表已为空")
            return
            
        if messagebox.askyesno("确认", "确定要清空所有文件吗？"):
            self.current_files.clear()
            self.file_listbox.delete(0, tk.END)
            messagebox.showinfo("成功", "已清除所有文件")
    
    def function_not_implemented(self, function_name):
        """功能未实现提示"""
        messagebox.showinfo("开发中", f"{function_name}功能正在开发中，敬请期待！")
    
    def show_merge_tool(self):
        """显示PDF合并工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建合并工具实例
        self.merge_tool = PDFMergeTool(self.right_frame, self.current_files)
    
    def show_split_tool(self):
        """显示PDF分割工具"""
        # 清除右侧区域
        self.clear_right_frame()
    
        # 创建分割工具实例
        self.split_tool = PDFSplitTool(self.right_frame, self.current_files)
    
    def show_compress_tool(self):
        """显示PDF压缩工具"""
        # 清除右侧区域
        self.clear_right_frame()
    
        # 创建压缩工具实例
        self.compress_tool = PDFCompressTool(self.right_frame, self.current_files)
    
    def show_encrypt_tool(self):
        """显示PDF加密/解密工具"""
        # 清除右侧区域
        self.clear_right_frame()
    
        # 创建加密解密工具实例
        self.encrypt_decrypt_tool = PDFEncryptDecryptTool(self.right_frame, self.current_files)

    def show_pdf_to_word_tool(self):
        """显示PDF转Word工具"""
        # 清除右侧区域
        self.clear_right_frame()
    
        # 创建PDF转Word工具实例
        self.pdf_to_word_tool = PDFToWordTool(self.right_frame, self.current_files)
    
    def show_pdf_to_image_tool(self):
        """显示PDF转图片工具"""
        # 清除右侧区域
        self.clear_right_frame()
    
        # 创建PDF与图片互转工具实例
        self.pdf_image_converter = PDFImageConverterTool(self.right_frame, self.current_files)
    
    def show_pdf_to_text_tool(self):
        """显示PDF转文本工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建PDF转文本工具实例
        self.pdf_to_text_tool = PDFToTextTool(self.right_frame, self.current_files)
    
    def show_word_to_pdf_tool(self):
        """显示Word转PDF工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建Word转PDF工具实例
        self.word_to_pdf_tool = WordToPDFTool(self.right_frame, self.current_files)
    
    def show_excel_to_pdf_tool(self):
        """显示Excel转PDF工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建Excel转PDF工具实例
        self.excel_to_pdf_tool = ExcelToPDFTool(self.right_frame, self.current_files)
    
    def show_ppt_to_pdf_tool(self):
        """显示PPT转PDF工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建PPT转PDF工具实例
        self.ppt_to_pdf_tool = PPTToPDFTool(self.right_frame, self.current_files)
    
    def show_watermark_tool(self):
        """显示PDF水印工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建PDF水印工具实例
        self.watermark_tool = PDFWatermarkTool(self.right_frame, self.current_files)
    
    def show_rotate_tool(self):
        """显示PDF旋转工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建PDF旋转工具实例
        self.rotate_tool = PDFRotateTool(self.right_frame, self.current_files)
    
    def show_header_footer_tool(self):
        """显示PDF页眉页脚工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建PDF页眉页脚工具实例
        self.header_footer_tool = PDFHeaderFooterTool(self.right_frame, self.current_files)
    
    def show_bookmark_tool(self):
        """显示PDF书签工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建PDF书签工具实例
        self.bookmark_tool = PDFBookmarkTool(self.right_frame, self.current_files)
    
    def show_annotation_tool(self):
        """显示PDF注释工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建PDF注释工具实例
        self.annotation_tool = PDFAnnotationTool(self.right_frame, self.current_files)
    
    def show_form_tool(self):
        """显示PDF表单工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建PDF表单工具实例
        self.form_tool = PDFFormTool(self.right_frame, self.current_files)
    
    def show_compare_tool(self):
        """显示PDF比较工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建PDF比较工具实例
        self.compare_tool = PDFCompareTool(self.right_frame, self.current_files)
    
    def show_ocr_tool(self):
        """显示OCR识别工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建OCR识别工具实例
        self.ocr_tool = PDFOCRTool(self.right_frame, self.current_files)
    
    def show_batch_tool(self):
        """显示批量处理工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建批量处理工具实例
        self.batch_tool = PDFBatchTool(self.right_frame, self.current_files)
    
    def show_optimize_tool(self):
        """显示PDF优化工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建PDF优化工具实例
        self.optimize_tool = PDFOptimizeTool(self.right_frame, self.current_files)
    
    def show_signature_tool(self):
        """显示PDF签名工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建PDF签名工具实例
        self.signature_tool = PDFSignatureTool(self.right_frame, self.current_files)
    
    def show_metadata_tool(self):
        """显示元数据编辑工具"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 创建元数据编辑工具实例
        self.metadata_tool = PDFMetadataTool(self.right_frame, self.current_files)
    
    def clear_right_frame(self):
        """清除右侧区域的所有内容"""
        for widget in self.right_frame.winfo_children():
            widget.destroy()
    
    def show_function_frame(self, frame_key):
        """显示指定的功能框架"""
        # 清除右侧区域
        self.clear_right_frame()
        
        # 重新创建功能框架
        if frame_key == "pdf_to_text":
            self.function_frames[frame_key] = self.create_function_frame("PDF转文本工具", "从PDF文件中提取文本内容")
        elif frame_key == "word_to_pdf":
            self.function_frames[frame_key] = self.create_function_frame("Word转PDF工具", "将Word文档转换为PDF格式")
        elif frame_key == "excel_to_pdf":
            self.function_frames[frame_key] = self.create_function_frame("Excel转PDF工具", "将Excel表格转换为PDF格式")
        elif frame_key == "ppt_to_pdf":
            self.function_frames[frame_key] = self.create_function_frame("PPT转PDF工具", "将PowerPoint演示文稿转换为PDF格式")
        elif frame_key == "watermark":
            self.function_frames[frame_key] = self.create_function_frame("PDF水印工具", "为PDF文件添加文字或图片水印")
        elif frame_key == "rotate":
            self.function_frames[frame_key] = self.create_function_frame("PDF旋转工具", "旋转PDF页面方向")
        elif frame_key == "header_footer":
            self.function_frames[frame_key] = self.create_function_frame("PDF页眉页脚工具", "为PDF文件添加页眉和页脚")
        elif frame_key == "bookmark":
            self.function_frames[frame_key] = self.create_function_frame("PDF书签工具", "为PDF文件添加或编辑书签")
        elif frame_key == "annotation":
            self.function_frames[frame_key] = self.create_function_frame("PDF注释工具", "为PDF文件添加注释、高亮和下划线")
        elif frame_key == "form":
            self.function_frames[frame_key] = self.create_function_frame("PDF表单工具", "填写或创建PDF表单")
        elif frame_key == "compare":
            self.function_frames[frame_key] = self.create_function_frame("PDF比较工具", "比较两个PDF文件的差异")
        elif frame_key == "ocr":
            self.function_frames[frame_key] = self.create_function_frame("OCR识别工具", "识别扫描PDF中的文字内容")
        elif frame_key == "batch":
            self.function_frames[frame_key] = self.create_function_frame("批量处理工具", "批量处理多个PDF文件")
        elif frame_key == "optimize":
            self.function_frames[frame_key] = self.create_function_frame("PDF优化工具", "优化PDF文件结构和性能")
        elif frame_key == "signature":
            self.function_frames[frame_key] = self.create_function_frame("PDF签名工具", "为PDF文件添加数字签名")
        elif frame_key == "metadata":
            self.function_frames[frame_key] = self.create_function_frame("元数据编辑工具", "编辑PDF文件的元数据信息")
        
        # 显示选中的功能框架
        if frame_key in self.function_frames:
            self.function_frames[frame_key].pack(fill=tk.BOTH, expand=True)

def main():
    root = tk.Tk()
    app = PDFToolbox(root)
    root.mainloop()

if __name__ == "__main__":
    main()
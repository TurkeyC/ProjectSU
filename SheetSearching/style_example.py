# 安装必要库：pip install customtkinter Pillow

import customtkinter as ctk
from PIL import Image
import os


class ModernUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 基础配置
        self._set_appearance_mode("Dark")  # 深色主题
        self.title("AI助手 v2.0")
        self.geometry("1200x800")
        self.minsize(1000, 700)

        # 加载图标资源
        self.icon_path = {
            "chat": self._load_image("chat.jpg", (24, 24)),
            "settings": self._load_image("chat.jpg", (24, 24)),
            "analytics": self._load_image("chat.jpg", (24, 24))
        }

        # 构建界面
        self._create_sidebar()
        self._create_main_content()
        self._create_bottom_bar()

    def _load_image(self, path, size):
        """加载并缩放图片"""
        return ctk.CTkImage(
            light_image=Image.open(path),
            dark_image=Image.open(path),
            size=size
        )

    def _create_sidebar(self):
        """左侧导航栏"""
        self.sidebar = ctk.CTkFrame(self, width=240, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")

        # 导航按钮
        nav_items = [
            ("对话模式", "chat", self.show_chat),
            ("数据分析", "analytics", self.show_analytics),
            ("系统设置", "settings", self.show_settings)
        ]

        for text, icon, command in nav_items:
            btn = ctk.CTkButton(
                self.sidebar,
                text=text,
                image=self.icon_path[icon],
                command=command,
                fg_color="transparent",
                hover_color=("#2B2B2B", "#3A3A3A"),
                anchor="w",
                height=48
            )
            btn.pack(fill="x", padx=10, pady=5)

        # 主题切换
        self.theme_switch = ctk.CTkSwitch(
            self.sidebar,
            text="深色模式",
            command=self._toggle_theme,
            progress_color="#2FA572"
        )
        self.theme_switch.pack(side="bottom", pady=20)

    def _create_main_content(self):
        """主内容区域"""
        self.main_frame = ctk.CTkFrame(self, corner_radius=8)
        self.main_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        # 选项卡视图
        self.tab_view = ctk.CTkTabview(self.main_frame)
        self.tab_view.pack(fill="both", expand=True)

        # 添加选项卡
        self.chat_tab = self.tab_view.add("AI对话")
        self.analytics_tab = self.tab_view.add("数据分析")
        self.settings_tab = self.tab_view.add("系统设置")

        # 初始化聊天界面
        self._init_chat_interface()

    def _init_chat_interface(self):
        """聊天界面组件"""
        # 聊天记录区
        self.chat_history = ctk.CTkScrollableFrame(self.chat_tab)
        self.chat_history.pack(fill="both", expand=True, padx=10, pady=10)

        # 输入区域
        input_frame = ctk.CTkFrame(self.chat_tab, height=120)
        input_frame.pack(fill="x", padx=10, pady=10)

        self.input_text = ctk.CTkEntry(
            input_frame,
            placeholder_text="输入您的问题...",
            height=80
        )
        self.input_text.pack(fill="x", padx=10, pady=10)

        send_btn = ctk.CTkButton(
            input_frame,
            text="发送",
            command=self.send_message,
            width=100,
            fg_color="#2FA572",
            hover_color="#1E7A5A"
        )
        send_btn.pack(side="right", padx=10)

    def _create_bottom_bar(self):
        """底部状态栏"""
        self.status_bar = ctk.CTkFrame(self, height=32)
        self.status_bar.pack(side="bottom", fill="x")

        self.status_label = ctk.CTkLabel(
            self.status_bar,
            text="就绪",
            anchor="w",
            font=("Segoe UI", 12)
        )
        self.status_label.pack(fill="x", padx=20)

    def _toggle_theme(self):
        """切换主题模式"""
        current = ctk.get_appearance_mode()
        new_mode = "Dark" if current == "Light" else "Light"
        ctk.set_appearance_mode(new_mode)
        self.theme_switch.configure(text=f"{new_mode}模式")

    def send_message(self):
        """处理消息发送"""
        message = self.input_text.get()
        if message:
            self._add_message("user", message)
            self.input_text.delete(0, "end")
            # 这里可以添加AI处理逻辑

    def _add_message(self, role, content):
        """添加聊天消息"""
        message_frame = ctk.CTkFrame(
            self.chat_history,
            fg_color=("#E0E0E0", "#2B2B2B") if role == "user" else ("#F0F0F0", "#3A3A3A")
        )
        message_frame.pack(fill="x", pady=5)

        label = ctk.CTkLabel(
            message_frame,
            text=content,
            wraplength=600,
            justify="left",
            font=("Segoe UI", 14)
        )
        label.pack(padx=20, pady=10)

    def show_chat(self):
        self.tab_view.set("AI对话")

    def show_analytics(self):
        self.tab_view.set("数据分析")

    def show_settings(self):
        self.tab_view.set("系统设置")


if __name__ == "__main__":
    app = ModernUI()
    app.mainloop()
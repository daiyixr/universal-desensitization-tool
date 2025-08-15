import os
import sys
import re
import tempfile
import requests
import webbrowser
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QTextEdit, QPushButton, QFileDialog, QMessageBox, 
    QGroupBox, QScrollArea, QInputDialog, QDialog, 
    QLineEdit, QDialogButtonBox, QComboBox, QProgressBar,
    QMenu, QAction, QTabWidget, QTableWidget, QTableWidgetItem, QCheckBox
)
from PyQt5.QtGui import (QIcon, QColor, QPalette, QLinearGradient, 
                         QBrush, QFont, QPixmap, QPainter)
from PyQt5.QtCore import Qt
from dataclasses import dataclass
from typing import List, Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet
    from openpyxl.cell.cell import Cell
    from openpyxl.workbook.workbook import Workbook

@dataclass
class RedactionRule:
    """脱敏规则数据类"""
    rule_id: str
    name: str
    pattern: str  # 正则表达式模式
    example: str
    is_active: bool = True
    regex: str = ""  # 存储编译后的正则表达式
    marker_char: str = ""  # 标记字符，用于全局搜索替换

class RuleEngine:
    """规则引擎核心类"""
    
    def __init__(self):
        self.rules: List[RedactionRule] = []
        
    def add_rule(self, rule: RedactionRule) -> None:
        """添加新规则"""
        self.rules.append(rule)
        
    def get_active_rules(self) -> List[RedactionRule]:
        """获取所有激活的规则"""
        return [r for r in self.rules if r.is_active]
        
    def load_default_rules(self) -> None:
        """加载默认规则集"""
        # 使用内置的完整规则集，默认只激活姓名规则
        self.rules = [
            # 身份信息类（高风险）
            RedactionRule(
                rule_id="name_rule",
                name="姓名",
                pattern="[\u4e00-\u9fa5]{2,4}",
                example="张三 → 张*",
                regex="[\u4e00-\u9fa5]{2,4}",
                marker_char="*",
                is_active=True  # 默认激活姓名规则
            ),
            RedactionRule(
                rule_id="id_card_rule",
                name="身份证号",
                pattern="[1-9]\\d{5}(18|19|20)\\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\\d|3[01])\\d{3}[0-9Xx]",
                example="123456199001011234 → 123***********1234",
                regex="[1-9]\\d{5}(18|19|20)\\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\\d|3[01])\\d{3}[0-9Xx]",
                marker_char="*",
                is_active=False  # 默认关闭其他规则
            ),
            RedactionRule(
                rule_id="passport_rule",
                name="护照号码",
                pattern="[EG]\\d{8}|[A-Z]\\d{7,8}",
                example="E12345678 → E1****678",
                regex="[EG]\\d{8}|[A-Z]\\d{7,8}",
                marker_char="*",
                is_active=False  # 默认关闭
            ),
            
            # 联系方式类（高风险）
            RedactionRule(
                rule_id="phone_rule",
                name="手机号码",
                pattern="1[3-9]\\d{9}",
                example="13812345678 → 138****5678",
                regex="1[3-9]\\d{9}",
                marker_char="*",
                is_active=False  # 默认关闭
            ),
            RedactionRule(
                rule_id="landline_rule",
                name="座机号码",
                pattern="0\\d{2,3}-?\\d{7,8}",
                example="010-12345678 → 010-****5678",
                regex="0\\d{2,3}-?\\d{7,8}",
                marker_char="*",
                is_active=False  # 默认关闭
            ),
            RedactionRule(
                rule_id="email_rule",
                name="邮箱地址",
                pattern="[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}",
                example="test@example.com → t***@example.com",
                regex="[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}",
                marker_char="*",
                is_active=False  # 默认关闭
            ),
            
            # 金融信息类（高风险）
            RedactionRule(
                rule_id="bank_card_rule",
                name="银行卡号",
                pattern="\\d{13,19}",
                example="6228480402564890018 → 6228****0018",
                regex="\\d{13,19}",
                marker_char="*",
                is_active=False
            ),
            
            # 地址信息类（中风险）
            RedactionRule(
                rule_id="address_rule",
                name="详细地址",
                pattern="[\u4e00-\u9fa5]{2,}(省|市|区|县|镇|街道|路|号|室|楼|层)[\u4e00-\u9fa5\\d\\-#]*",
                example="北京市朝阳区建国路1号 → 北京市朝阳区建******",
                regex="[\u4e00-\u9fa5]{2,}(省|市|区|县|镇|街道|路|号|室|楼|层)[\u4e00-\u9fa5\\d\\-#]*",
                marker_char="*",
                is_active=False
            ),
            
            # 车辆信息类（中风险）
            RedactionRule(
                rule_id="license_plate_rule",
                name="车牌号码",
                pattern="[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领][A-Z][A-Z0-9]{4}[A-Z0-9挂学警港澳]",
                example="京A12345 → 京A***45",
                regex="[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领][A-Z][A-Z0-9]{4}[A-Z0-9挂学警港澳]",
                marker_char="*",
                is_active=False
            ),
            
            # 企业信息类（中风险）
            RedactionRule(
                rule_id="organization_code_rule",
                name="组织机构代码",
                pattern="[A-Z0-9]{8}-[A-Z0-9]|[A-Z0-9]{9}",
                example="12345678-9 → 123***8-9",
                regex="[A-Z0-9]{8}-[A-Z0-9]|[A-Z0-9]{9}",
                marker_char="*",
                is_active=False
            ),
            RedactionRule(
                rule_id="tax_id_rule",
                name="纳税人识别号",
                pattern="\\d{15}|\\d{17}[0-9X]|\\d{18}|\\d{20}",
                example="123456789012345 → 1234*******2345",
                regex="\\d{15}|\\d{17}[0-9X]|\\d{18}|\\d{20}",
                marker_char="*",
                is_active=False
            ),
            
            # 内部标识类（低风险）
            RedactionRule(
                rule_id="employee_id_rule",
                name="员工工号",
                pattern="[A-Z0-9]{4,10}",
                example="EMP001234 → EMP***234",
                regex="[A-Z0-9]{4,10}",
                marker_char="*",
                is_active=False
            ),
            
            # 自定义字段类（用户可配置）
            RedactionRule(
                rule_id="custom_field_rule",
                name="自定义字段",
                pattern=".*",
                example="任意字段 → 任意*",
                regex=".*",
                marker_char="*",
                is_active=False
            )
        ]
            
    def verify_rule_examples(self):
        """验证所有规则的example和实际脱敏效果是否一致"""
        print("开始验证规则示例...")
        for rule in self.rules:
            if " → " in rule.example:
                parts = rule.example.split(" → ")
                if len(parts) == 2:
                    input_text = parts[0]
                    expected_output = parts[1]
                    actual_output = self.apply_redaction_rule(rule, input_text, None)
                    
                    if actual_output != expected_output:
                        print(f"规则 {rule.name} 不一致:")
                        print(f"  输入: {input_text}")
                        print(f"  期望: {expected_output}")
                        print(f"  实际: {actual_output}")
                    else:
                        print(f"规则 {rule.name} ✓")
        print("验证完成")

    def apply_redaction_rule(self, rule: RedactionRule, text: str, custom_list=None) -> str:
        """应用脱敏规则到文本"""
        if not rule.regex or not text:
            return text
            
        try:
            import re
            
            # 姓名规则特殊处理：只替换自定义名单中的姓名
            if rule.rule_id == "name_rule":
                result = text
                
                # 检查是否有自定义姓名列表
                if custom_list and isinstance(custom_list, list):
                    # 只对自定义名单中的姓名进行替换
                    for name in custom_list:
                        if name and name in result:
                            # 姓名脱敏：保留第一个字，其余用*替换
                            redacted_name = name[0] + "*" * (len(name) - 1)
                            result = result.replace(name, redacted_name)
                    return result
                else:
                    # 没有自定义名单时，不进行任何替换
                    return text
            
            # 自定义字段规则特殊处理：完全复用姓名规则的逻辑
            if rule.rule_id == "custom_field_rule":
                result = text
                
                # 检查是否有自定义字段列表
                if custom_list and isinstance(custom_list, list):
                    # 只对自定义字段列表中的内容进行替换
                    for field in custom_list:
                        if field and field in result:
                            # 自定义字段脱敏：保留第一个字符，其余用*替换（与姓名规则完全相同）
                            if len(field) > 0:
                                redacted_field = field[0] + "*" * (len(field) - 1)
                                result = result.replace(field, redacted_field)
                    return result
                else:
                    # 没有自定义字段列表时，不进行任何替换
                    return text
            
            # 其他规则保持原有逻辑
            # 查找所有匹配的内容
            matches = re.findall(rule.regex, text)
            if not matches:
                return text
                
            result = text
            
            # 处理复杂匹配（可能包含组的情况）
            processed_matches = []
            for match in matches:
                if isinstance(match, tuple):
                    # 如果是元组（来自分组），取非空的组
                    actual_match = next((group for group in match if group), "")
                else:
                    actual_match = match
                
                if actual_match:
                    processed_matches.append(actual_match)
            
            # 如果没有有效匹配，使用原始匹配逻辑
            if not processed_matches:
                processed_matches = [m for m in matches if m]
            
            for match in processed_matches:
                # 根据规则ID或名称应用不同的脱敏策略
                if rule.rule_id == "id_card_rule" and len(match) == 18:
                    # 身份证脱敏：保留前3位和后4位
                    redacted = match[:3] + "*" * 11 + match[-4:]
                elif rule.rule_id == "phone_rule" and len(match) == 11:
                    # 手机号脱敏：保留前3位和后4位
                    redacted = match[:3] + "****" + match[-4:]
                elif rule.rule_id == "landline_rule":
                    # 座机号脱敏：保留区号和后4位，中间用*替换
                    if "-" in match:
                        parts = match.split("-")
                        if len(parts[1]) > 4:
                            redacted = parts[0] + "-" + "*" * (len(parts[1]) - 4) + parts[1][-4:]
                        else:
                            redacted = parts[0] + "-" + "*" * len(parts[1])
                    else:
                        if len(match) > 8:
                            redacted = match[:4] + "*" * (len(match) - 8) + match[-4:]
                        else:
                            redacted = match[:4] + "*" * (len(match) - 4)
                elif rule.rule_id == "email_rule":
                    # 邮箱脱敏：保留第一个字符和@后的内容
                    at_index = match.find('@')
                    if at_index > 0:
                        redacted = match[0] + "*" * (at_index - 1) + match[at_index:]
                    else:
                        redacted = match
                elif rule.rule_id == "address_rule":
                    # 地址脱敏：保留前6个字符，其余用*替换
                    if len(match) > 6:
                        redacted = match[:6] + "*" * (len(match) - 6)
                    else:
                        redacted = "*" * len(match)
                elif rule.rule_id == "bank_card_rule":
                    # 银行卡号脱敏：保留前4位和后4位
                    if len(match) > 8:
                        redacted = match[:4] + "*" * (len(match) - 8) + match[-4:]
                    else:
                        redacted = "*" * len(match)
                elif rule.rule_id == "license_plate_rule":
                    # 车牌号脱敏：保留省份+字母和最后2位，中间用*替换
                    if len(match) >= 7:
                        redacted = match[:2] + "*" * (len(match) - 4) + match[-2:]
                    else:
                        redacted = "*" * len(match)
                elif rule.rule_id == "passport_rule":
                    # 护照号脱敏：保留前2位和后3位
                    if len(match) > 5:
                        redacted = match[:2] + "*" * (len(match) - 5) + match[-3:]
                    else:
                        redacted = "*" * len(match)
                elif rule.rule_id == "organization_code_rule":
                    # 组织机构代码脱敏：保留前3位和后2位（包括-后的部分）
                    if "-" in match:
                        parts = match.split("-")
                        if len(parts[0]) > 3:
                            redacted = parts[0][:3] + "*" * (len(parts[0]) - 4) + parts[0][-1] + "-" + parts[1]
                        else:
                            redacted = match
                    else:
                        if len(match) > 5:
                            redacted = match[:3] + "*" * (len(match) - 5) + match[-2:]
                        else:
                            redacted = "*" * len(match)
                elif rule.rule_id == "tax_id_rule":
                    # 纳税人识别号脱敏：保留前4位和后4位
                    if len(match) > 8:
                        redacted = match[:4] + "*" * (len(match) - 8) + match[-4:]
                    else:
                        redacted = "*" * len(match)
                elif rule.rule_id == "employee_id_rule":
                    # 员工工号脱敏：保留前3位和后3位，中间用*替换
                    if len(match) > 6:
                        redacted = match[:3] + "*" * (len(match) - 6) + match[-3:]
                    elif len(match) > 3:
                        redacted = match[:3] + "*" * (len(match) - 3)
                    else:
                        redacted = "*" * len(match)
                else:
                    # 改进的默认脱敏方式：使用智能脱敏逻辑
                    # 调用智能脱敏函数，避免全星号替换
                    redacted = self.smart_redact_for_rule_engine(match)
                
                result = result.replace(match, redacted)
            return result
        except Exception:
            return text
    
    def smart_redact_for_rule_engine(self, text):
        """规则引擎专用的智能脱敏函数"""
        import re
        
        # 检测文本类型并应用相应脱敏规则
        text = text.strip()
        
        # 身份证号（18位）
        if re.match(r'^\d{18}$', text):
            return text[:3] + "*" * 11 + text[-4:]
        
        # 手机号（11位数字）
        elif re.match(r'^1[3-9]\d{9}$', text):
            return text[:3] + "****" + text[-4:]
        
        # 邮箱地址
        elif '@' in text and re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', text):
            at_index = text.find('@')
            return text[0] + "*" * (at_index - 1) + text[at_index:]
        
        # 中文姓名（2-4个汉字）
        elif re.match(r'^[\u4e00-\u9fa5]{2,4}$', text):
            return text[0] + "*" * (len(text) - 1)
        
        # 银行卡号（16-19位数字）
        elif re.match(r'^\d{16,19}$', text):
            return text[:4] + "*" * (len(text) - 8) + text[-4:]
        
        # 座机号码（带区号）
        elif re.match(r'^0\d{2,3}-?\d{7,8}$', text):
            if '-' in text:
                parts = text.split('-')
                return parts[0] + "-****" + parts[1][-4:]
            else:
                return text[:4] + "****" + text[-4:]
        
        # 车牌号
        elif re.match(r'^[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领][A-Z]\d{5}$', text):
            return text[:2] + "***" + text[-2:]
        
        # 护照号
        elif re.match(r'^[A-Z]\d{8}$', text):
            return text[:2] + "****" + text[-3:]
        
        # 如果不匹配任何模式，采用保守的脱敏策略
        else:
            # 对于短文本（小于等于3个字符），保留第一个字符
            if len(text) <= 3:
                return text[0] + "*" * (len(text) - 1)
            # 对于长文本，保留前后各1个字符
            elif len(text) <= 10:
                return text[0] + "*" * (len(text) - 2) + text[-1]
            # 对于很长的文本，保留前后各2个字符
            else:
                return text[:2] + "*" * (len(text) - 4) + text[-2:]

class UniversalRedactionTool(QMainWindow):
    def check_update(self):
        """检查软件更新"""
        import webbrowser
        url = "https://gitee.com/daiyixr/universal-desensitization-tool/raw/master/latest_version.json"
        try:
            info = requests.get(url, timeout=10).json()
            latest = info.get("version", "")
            if latest and latest != self.version:
                # 检测到新版本，直接打开百度网盘下载页面
                baidu_url = "https://pan.baidu.com/s/1_eiYyKkYYMZa3ExVkLt3rg?pwd=muxz"  # 替换为你的真实链接
                webbrowser.open(baidu_url)
                # 同时提示提取码
                QMessageBox.information(
                    self, 
                    "发现新版本", 
                    f"当前版本：{self.version}\n最新版本：{latest}\n\n更新内容：{info.get('desc', '')}\n\n已自动打开下载页面\n提取码：muxz（永久有效）"  # 替换为你的真实提取码
                )
            else:
                QMessageBox.information(self, "检查更新", "当前已是最新版本！")
        except Exception as e:
            QMessageBox.warning(self, "检查更新失败", f"检查更新时发生错误：{str(e)}")

    def show_name_redact_dialog(self):
        """弹窗：用户自定义输入需脱敏的姓名列表，支持文本/表格粘贴，自动识别并反馈未识别项"""
        dialog = QDialog(self)
        dialog.setWindowTitle("自定义姓名脱敏")
        dialog.setWindowIcon(self.windowIcon())
        dialog.resize(500, 400)
        dialog.setModal(True)
        layout = QVBoxLayout()
        
        # 说明标签
        instruction_label = QLabel("请粘贴所有需要脱敏的姓名（每行一个，支持表格/文本粘贴）：")
        instruction_label.setStyleSheet("font-weight: bold; color: #2c3e50; margin-bottom: 10px;")
        layout.addWidget(instruction_label)
        
        # 姓名输入框
        name_edit = QTextEdit()
        name_edit.setPlaceholderText("如：\n张三\n李四\n王五\n\n支持批量粘贴，一行一个姓名")
        name_edit.setMinimumHeight(150)
        layout.addWidget(name_edit)
        
        # 结果显示标签
        result_label = QLabel("")
        result_label.setStyleSheet("color: #27ae60; font-size: 10pt; background-color: #f8f9fa; padding: 8px; border-radius: 4px; margin: 10px 0;")
        result_label.setWordWrap(True)
        layout.addWidget(result_label)
        
        # 存储有效姓名的变量
        valid_names = []
        
        def on_confirm():
            nonlocal valid_names
            raw_text = name_edit.toPlainText().strip()
            if not raw_text:
                QMessageBox.warning(dialog, "提示", "请输入姓名列表！")
                return
            
            # 处理姓名列表
            names = [n.strip() for n in raw_text.splitlines() if n.strip()]
            # 简单中文姓名识别（2-4位汉字）
            valid = [n for n in names if re.match(r"^[\u4e00-\u9fa5]{2,4}$", n)]
            invalid = [n for n in names if n not in valid]
            
            valid_names = valid  # 保存有效姓名
            
            msg = f"✅ 识别成功 {len(valid)} 个姓名"
            if valid:
                msg += f"：\n{', '.join(valid[:10])}"
                if len(valid) > 10:
                    msg += f" ... 等{len(valid)}个"
            
            if invalid:
                msg += f"\n\n❌ 未识别（请检查格式）{len(invalid)}个：\n{', '.join(invalid[:5])}"
                if len(invalid) > 5:
                    msg += f" ... 等{len(invalid)}个"
            
            result_label.setText(msg)
            
            if valid:
                # 将有效姓名添加到姓名规则中
                self.update_name_rule_with_custom_names(valid)
                save_btn.setText("✅ 已识别")
                save_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #27ae60;
                        color: white;
                        border: none;
                        border-radius: 5px;
                        padding: 10px 20px;
                        font-size: 14px;
                        font-weight: bold;
                    }
                    QPushButton:hover {
                        background-color: #229954;
                    }
                """)
        
        def on_save_and_close():
            if valid_names:
                QMessageBox.information(dialog, "保存成功", f"已成功保存 {len(valid_names)} 个自定义姓名到脱敏规则中。\n\n在进行脱敏处理时，这些姓名将被自动识别和脱敏。")
                dialog.accept()
            else:
                # 如果没有已识别的姓名，先尝试识别
                on_confirm()
                if valid_names:
                    QMessageBox.information(dialog, "保存成功", f"已成功保存 {len(valid_names)} 个自定义姓名到脱敏规则中。\n\n在进行脱敏处理时，这些姓名将被自动识别和脱敏。")
                    dialog.accept()
                else:
                    QMessageBox.warning(dialog, "提示", "请先输入并识别姓名后再保存！")
        
        def on_cancel_name():
            # 取消操作：仅关闭窗口
            dialog.reject()
        
        # 按钮区域
        btn_layout = QHBoxLayout()
        
        save_btn = QPushButton("确定并识别")
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        save_btn.clicked.connect(on_confirm)
        
        close_btn = QPushButton("取消")
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #7f8c8d;
            }
        """)
        close_btn.clicked.connect(on_cancel_name)
        
        save_and_close_btn = QPushButton("保存并关闭")
        save_and_close_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        save_and_close_btn.clicked.connect(on_save_and_close)
        
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(save_and_close_btn)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)
        
        dialog.setLayout(layout)
        dialog.exec_()

    def show_custom_field_redact_dialog(self):
        """弹窗：用户自定义输入需脱敏的任意字段列表，支持文本/表格粘贴，无格式限制"""
        dialog = QDialog(self)
        dialog.setWindowTitle("自定义字段脱敏")
        dialog.setWindowIcon(self.windowIcon())
        dialog.resize(500, 400)
        dialog.setModal(True)
        layout = QVBoxLayout()
        
        # 说明标签
        instruction_label = QLabel("请粘贴所有需要脱敏的字段内容（每行一个，支持任意字符）：")
        instruction_label.setStyleSheet("font-weight: bold; color: #2c3e50; margin-bottom: 10px;")
        layout.addWidget(instruction_label)
        
        # 字段输入框
        field_edit = QTextEdit()
        field_edit.setPlaceholderText("如：\n公司A\n部门01\nABC123\n项目-X\n\n支持批量粘贴，一行一个字段\n支持中文、英文、数字、符号等任意字符")
        field_edit.setMinimumHeight(150)
        layout.addWidget(field_edit)
        
        # 结果显示标签
        result_label = QLabel("")
        result_label.setStyleSheet("color: #27ae60; font-size: 10pt; background-color: #f8f9fa; padding: 8px; border-radius: 4px; margin: 10px 0;")
        result_label.setWordWrap(True)
        layout.addWidget(result_label)
        
        # 存储有效字段的变量
        valid_fields = []
        
        def on_confirm():
            nonlocal valid_fields
            raw_text = field_edit.toPlainText().strip()
            if not raw_text:
                QMessageBox.warning(dialog, "提示", "请输入字段列表！")
                return
            
            # 处理字段列表
            fields = [f.strip() for f in raw_text.splitlines() if f.strip()]
            # 对于自定义字段，接受任何非空内容
            valid = [f for f in fields if f]
            
            valid_fields = valid  # 保存有效字段
            
            msg = f"✅ 识别成功 {len(valid)} 个字段"
            if valid:
                msg += f"：\n{', '.join(valid[:10])}"
                if len(valid) > 10:
                    msg += f" ... 等{len(valid)}个"
            
            result_label.setText(msg)
            
            if valid:
                # 将有效字段添加到自定义字段规则中
                self.update_custom_field_rule_with_fields(valid)
                save_btn.setText("✅ 已识别")
                save_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #27ae60;
                        color: white;
                        border: none;
                        border-radius: 5px;
                        padding: 10px 20px;
                        font-size: 14px;
                        font-weight: bold;
                    }
                    QPushButton:hover {
                        background-color: #229954;
                    }
                """)
        
        def on_save_and_close():
            if valid_fields:
                QMessageBox.information(dialog, "保存成功", f"已成功保存 {len(valid_fields)} 个自定义字段到脱敏规则中。\n\n在进行脱敏处理时，这些字段将被自动识别和脱敏。")
                dialog.accept()
            else:
                # 如果没有已识别的字段，先尝试识别
                on_confirm()
                if valid_fields:
                    QMessageBox.information(dialog, "保存成功", f"已成功保存 {len(valid_fields)} 个自定义字段到脱敏规则中。\n\n在进行脱敏处理时，这些字段将被自动识别和脱敏。")
                    dialog.accept()
                else:
                    QMessageBox.warning(dialog, "提示", "请先输入并识别字段后再保存！")
        
        # 按钮区域
        btn_layout = QHBoxLayout()
        
        save_btn = QPushButton("确定并识别")
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        save_btn.clicked.connect(on_confirm)
        
        def on_cancel():
            # 取消操作：仅关闭窗口
            dialog.reject()
        
        close_btn = QPushButton("取消")
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #7f8c8d;
            }
        """)
        close_btn.clicked.connect(on_cancel)
        
        save_and_close_btn = QPushButton("保存并关闭")
        save_and_close_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        save_and_close_btn.clicked.connect(on_save_and_close)
        
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(save_and_close_btn)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)
        
        dialog.setLayout(layout)
        dialog.exec_()

    def update_name_rule_with_custom_names(self, custom_names, save_to_file=True):
        """将自定义姓名更新到姓名脱敏规则中"""
        if not custom_names:
            return
        
        # 找到姓名规则
        name_rule = None
        for rule in self.rule_engine.rules:
            if rule.rule_id == "name_rule" or rule.name == "姓名":
                name_rule = rule
                break
        
        if name_rule:
            # 创建自定义姓名的正则表达式
            # 转义特殊字符，然后用 | 连接
            escaped_names = [re.escape(name) for name in custom_names]
            custom_pattern = "|".join(escaped_names)
            
            # 更新规则的匹配模式，同时保留原有的中文姓名通用匹配
            original_pattern = "[\u4e00-\u9fa5]{2,4}"  # 原有的通用中文姓名匹配
            combined_pattern = f"({custom_pattern})|({original_pattern})"
            
            name_rule.pattern = combined_pattern
            name_rule.regex = combined_pattern
            name_rule.is_active = True  # 确保规则被激活
            
            # 更新示例文本
            example_names = custom_names[:3]  # 取前3个作为示例
            if len(custom_names) > 3:
                example_text = f"{example_names[0]} → {example_names[0][0]}*，{example_names[1]} → {example_names[1][0]}*... 等{len(custom_names)}个"
            else:
                example_text = "，".join([f"{name} → {name[0]}*" for name in example_names])
            name_rule.example = example_text
            
            # 保存自定义姓名列表（用于后续显示）
            self.custom_names = custom_names

            # 只在需要时保存到文件（避免加载时重复保存）
            if save_to_file:
                self.save_unified_custom_rules(custom_names=custom_names)
            print(f"已更新姓名规则，包含 {len(custom_names)} 个自定义姓名")
        else:
            QMessageBox.warning(self, "错误", "未找到姓名脱敏规则，无法添加自定义姓名")

    def save_unified_custom_rules(self, custom_names=None, custom_fields=None):
        """统一保存所有自定义规则到一个JSON文件中（每天一个文件）"""
        import os, json, datetime
        rules_dir = os.path.join(os.path.dirname(__file__), "user_custom_rules")
        if not os.path.exists(rules_dir):
            os.makedirs(rules_dir)
        
        date_str = datetime.datetime.now().strftime("%Y%m%d")
        file_path = os.path.join(rules_dir, f"自定义规则{date_str}.json")
        
        # 加载现有数据（如果文件存在）
        existing_data = {}
        if os.path.exists(file_path):
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    existing_data = json.load(f)
            except Exception as e:
                print(f"读取现有规则文件失败: {e}")
        
        # 更新数据
        if custom_names is not None:
            existing_data["custom_names"] = custom_names
        if custom_fields is not None:
            existing_data["custom_fields"] = custom_fields
        
        # 保存更新后的数据
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(existing_data, f, ensure_ascii=False, indent=2)
            print(f"已保存自定义规则到: {file_path}")
        except Exception as e:
            print(f"保存自定义规则失败: {e}")

    def load_unified_custom_rules(self):
        """从统一的自定义规则文件中加载所有规则"""
        import os, json, glob
        rules_dir = os.path.join(os.path.dirname(__file__), "user_custom_rules")
        if not os.path.exists(rules_dir):
            return
        
        files = glob.glob(os.path.join(rules_dir, "自定义规则*.json"))
        if not files:
            return
            
        # 获取最新的文件
        latest_file = max(files, key=os.path.getmtime)
        try:
            with open(latest_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            # 加载自定义姓名
            custom_names = data.get("custom_names", [])
            if custom_names:
                self.update_name_rule_with_custom_names(custom_names, save_to_file=False)
                print(f"已加载 {len(custom_names)} 个自定义姓名")
            
            # 加载自定义字段
            custom_fields = data.get("custom_fields", [])
            if custom_fields:
                self.update_custom_field_rule_with_fields(custom_fields, save_to_file=False)
                print(f"已加载 {len(custom_fields)} 个自定义字段")
                
            if custom_names or custom_fields:
                print(f"已自动加载最新自定义规则：{latest_file}")
        except Exception as e:
            print(f"加载自定义规则失败: {e}")

    def update_custom_field_rule_with_fields(self, custom_fields, save_to_file=True):
        """将自定义字段更新到自定义字段脱敏规则中"""
        if not custom_fields:
            return
        
        # 找到自定义字段规则
        custom_field_rule = None
        for rule in self.rule_engine.rules:
            if rule.rule_id == "custom_field_rule" or rule.name == "自定义字段":
                custom_field_rule = rule
                break
        
        if custom_field_rule:
            # 创建自定义字段的正则表达式
            # 转义特殊字符，然后用 | 连接
            escaped_fields = [re.escape(field) for field in custom_fields]
            custom_pattern = "|".join(escaped_fields)
            
            # 更新规则的匹配模式
            custom_field_rule.pattern = custom_pattern
            custom_field_rule.regex = custom_pattern
            custom_field_rule.is_active = True  # 确保规则被激活
            
            # 更新示例文本
            example_fields = custom_fields[:3]  # 取前3个作为示例
            if len(custom_fields) > 3:
                example_text = f"{example_fields[0]} → {example_fields[0][0]}*，{example_fields[1]} → {example_fields[1][0]}*... 等{len(custom_fields)}个"
            else:
                example_text = "，".join([f"{field} → {field[0]}*" for field in example_fields if field])
            custom_field_rule.example = example_text
            
            # 保存自定义字段列表（用于后续显示）
            self.custom_fields = custom_fields

            # 只在需要时保存到文件（避免加载时重复保存）
            if save_to_file:
                self.save_unified_custom_rules(custom_fields=custom_fields)
            print(f"已更新自定义字段规则，包含 {len(custom_fields)} 个自定义字段")
        else:
            QMessageBox.warning(self, "错误", "未找到自定义字段脱敏规则，无法添加自定义字段")

    def load_latest_custom_names(self):
        """自动加载最新的自定义规则JSON文件并应用到规则引擎（保持兼容性）"""
        # 调用统一的加载方法
        self.load_unified_custom_rules()

    def closeEvent(self, a0):
        """程序退出时的处理"""
        # 检查是否有自定义规则需要清除
        has_custom_names = hasattr(self, 'custom_names') and self.custom_names
        has_custom_fields = hasattr(self, 'custom_fields') and self.custom_fields
        
        if has_custom_names or has_custom_fields:
            reply = QMessageBox.question(self, '确认', 
                                       '是否清除本次使用的自定义规则？\n'
                                       '选择"是"将清除内存中的自定义规则并删除JSON文件（不可恢复）\n'
                                       '选择"否"将保留规则到下次启动',
                                       QMessageBox.StandardButton.Yes,
                                       QMessageBox.StandardButton.No)
            
            if reply == QMessageBox.StandardButton.Yes:
                # 清除自定义规则
                if has_custom_names:
                    self.custom_names = []
                    # 重置姓名规则为默认状态
                    for rule in self.rule_engine.rules:
                        if rule.rule_id == "name_rule" or rule.name == "姓名":
                            rule.pattern = "[\u4e00-\u9fa5]{2,4}"
                            rule.regex = "[\u4e00-\u9fa5]{2,4}"
                            rule.example = "张三 → 张*，李四 → 李*"
                            break
                
                if has_custom_fields:
                    self.custom_fields = []
                    # 重置自定义字段规则为默认状态
                    for rule in self.rule_engine.rules:
                        if rule.rule_id == "custom_field_rule" or rule.name == "自定义字段":
                            rule.pattern = ""
                            rule.regex = ""
                            rule.example = "请先配置自定义字段"
                            rule.is_active = False
                            break
                
                # 删除自定义JSON文件
                import os, glob
                try:
                    rules_dir = os.path.join(os.path.dirname(__file__), "user_custom_rules")
                    if os.path.exists(rules_dir):
                        files = glob.glob(os.path.join(rules_dir, "自定义规则*.json"))
                        deleted_count = 0
                        for file_path in files:
                            try:
                                os.remove(file_path)
                                deleted_count += 1
                            except Exception as e:
                                print(f"删除文件失败 {file_path}: {e}")
                        
                        if deleted_count > 0:
                            print(f"已清除自定义规则及 {deleted_count} 个JSON文件")
                        else:
                            print("已清除自定义规则")
                except Exception as e:
                    print(f"清除JSON文件时出错: {e}")
        
        super().closeEvent(a0)

    def __init__(self):
        super().__init__()
        self.version = "2.1.5"  # 添加版本号属性
        self.setWindowTitle("📋通用脱敏工具")
        self.setWindowIcon(self.get_app_icon())
        self.setGeometry(200, 120, 800, 650)
        self.setup_ui()
        self.setup_styles()
        
        # 初始化文档对象
        self.current_word_doc = None
        
        # 初始化撤销历史记录
        self.text_redaction_history = []  # 文本脱敏历史记录
        self.word_redaction_history = []  # Word文档脱敏历史记录
        self.excel_redaction_history = []  # Excel脱敏历史记录
        
        # 初始化Excel格式存储
        self.excel_cell_formats = {}  # 存储每个单元格的原始格式信息
        self.original_excel_path = None  # 存储原始Excel文件路径
        
        # 自动加载最新的自定义规则
        self.load_latest_custom_names()

    def save_cell_format(self, cell, row, col):
        """保存单元格的格式信息"""
        try:
            # 保存单元格的关键格式属性
            format_info = {
                'font': {
                    'name': cell.font.name if cell.font and cell.font.name else None,
                    'size': cell.font.size if cell.font and cell.font.size else None,
                    'bold': cell.font.bold if cell.font else None,
                    'italic': cell.font.italic if cell.font else None,
                    'color': str(cell.font.color.rgb) if cell.font and cell.font.color and hasattr(cell.font.color, 'rgb') else None,
                },
                'fill': {
                    'fill_type': str(cell.fill.fill_type) if cell.fill else None,
                    'start_color': str(cell.fill.start_color.rgb) if cell.fill and cell.fill.start_color and hasattr(cell.fill.start_color, 'rgb') else None,
                },
                'border': {
                    'left': str(cell.border.left.style) if cell.border and cell.border.left else None,
                    'right': str(cell.border.right.style) if cell.border and cell.border.right else None,
                    'top': str(cell.border.top.style) if cell.border and cell.border.top else None,
                    'bottom': str(cell.border.bottom.style) if cell.border and cell.border.bottom else None,
                },
                'alignment': {
                    'horizontal': str(cell.alignment.horizontal) if cell.alignment and cell.alignment.horizontal else None,
                    'vertical': str(cell.alignment.vertical) if cell.alignment and cell.alignment.vertical else None,
                    'wrap_text': cell.alignment.wrap_text if cell.alignment else None,
                },
                'number_format': cell.number_format if hasattr(cell, 'number_format') else None,
            }
            
            # 使用(row, col)作为键存储格式信息
            self.excel_cell_formats[(row, col)] = format_info
        except Exception as e:
            # 如果保存格式失败，忽略错误但记录日志
            print(f"警告：保存单元格({row}, {col})格式失败: {str(e)}")

    def apply_cell_format(self, cell, row, col):
        """将保存的格式应用到单元格"""
        try:
            format_info = self.excel_cell_formats.get((row, col))
            if not format_info:
                return
            
            from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
            
            # 应用字体格式
            font_info = format_info.get('font', {})
            if any(font_info.values()):
                cell.font = Font(
                    name=font_info.get('name'),
                    size=font_info.get('size'),
                    bold=font_info.get('bold'),
                    italic=font_info.get('italic'),
                    color=font_info.get('color')
                )
            
            # 应用填充格式
            fill_info = format_info.get('fill', {})
            if any(fill_info.values()):
                cell.fill = PatternFill(
                    fill_type=fill_info.get('fill_type'),
                    start_color=fill_info.get('start_color')
                )
            
            # 应用边框格式
            border_info = format_info.get('border', {})
            if any(border_info.values()):
                cell.border = Border(
                    left=Side(style=border_info.get('left')),
                    right=Side(style=border_info.get('right')),
                    top=Side(style=border_info.get('top')),
                    bottom=Side(style=border_info.get('bottom'))
                )
            
            # 应用对齐格式
            alignment_info = format_info.get('alignment', {})
            if any(alignment_info.values()):
                cell.alignment = Alignment(
                    horizontal=alignment_info.get('horizontal'),
                    vertical=alignment_info.get('vertical'),
                    wrap_text=alignment_info.get('wrap_text')
                )
            
            # 应用数字格式
            number_format = format_info.get('number_format')
            if number_format:
                cell.number_format = number_format
                
        except Exception as e:
            print(f"警告：应用单元格({row}, {col})格式失败: {str(e)}")

    def get_app_icon(self):
        # 创建简单图标，避免sRGB配置文件警告
        pixmap = QPixmap(32, 32)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # 绘制简单的文档图标
        painter.setBrush(QColor(67, 97, 238))  # 使用RGB值而不是十六进制
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(6, 4, 20, 24, 4, 4)
        
        painter.setBrush(QColor(255, 255, 255))
        painter.drawRect(10, 8, 12, 12)
        
        painter.setBrush(QColor(67, 97, 238))
        painter.drawRect(10, 22, 12, 2)
        
        painter.end()
        return QIcon(pixmap)

    def setup_ui(self):
        # 主窗口布局
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # 标题区域
        title_label = QLabel("通用脱敏工具")
        title_label.setFont(QFont("Arial", 18, QFont.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        version_label = QLabel(f"版本: V2.1.5 | 2025 D&Ai ")
        version_label.setObjectName("version_label")
        self.version_label = version_label  # 保存为实例变量
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        version_label.setStyleSheet("color: #7f8c8d; font-size: 10pt;")
        main_layout.addWidget(version_label)

        # 文件操作区（前置）
        file_group = QGroupBox("📁 文件处理")
        file_layout = QVBoxLayout()
        
        # 处理模式选择
        mode_layout = QHBoxLayout()
        mode_label = QLabel("处理模式:")
        mode_label.setFont(QFont("Arial", 10, QFont.Bold))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems([
            "🎯 交互式脱敏（推荐）",
            "⚙️ 自动脱敏（规则模式）"
        ])
        self.mode_combo.setCurrentIndex(0)  # 默认选择交互式脱敏
        self.mode_combo.currentIndexChanged.connect(self.on_mode_changed)
        self.rule_config_btn = QPushButton("📋 配置脱敏规则")
        self.rule_config_btn.setStyleSheet("""
            QPushButton {
                font-size: 18px;
                font-weight: bold;
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        self.rule_config_btn.clicked.connect(self.show_rule_config_dialog)
        self.rule_config_btn.setVisible(False)  # 初始隐藏
        
        mode_layout.addWidget(mode_label)
        mode_layout.addWidget(self.mode_combo)
        mode_layout.addWidget(self.rule_config_btn)
        mode_layout.addStretch()
        file_layout.addLayout(mode_layout)
        
        # 模式说明标签
        self.mode_tip_label = QLabel("💡 交互式脱敏：选中文本或单元格后右键选择脱敏，精确控制每个内容")
        self.mode_tip_label.setStyleSheet("color: #27ae60; font-size: 9pt; background-color: #d5f4e6; padding: 8px; border-radius: 5px; border-left: 3px solid #27ae60;")
        self.mode_tip_label.setWordWrap(True)
        file_layout.addWidget(self.mode_tip_label)
        
        # 文件选择按钮
        file_btn_layout = QHBoxLayout()
        self.input_btn = QPushButton("📂 选择待脱敏文件")
        self.input_btn.setMinimumHeight(40)
        self.input_btn.setStyleSheet("""
            QPushButton {
                font-size: 18px;
                font-weight: bold;
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        self.input_btn.clicked.connect(self.select_input_file)
        
        self.output_btn = QPushButton("💾 设置输出路径") 
        self.output_btn.setMinimumHeight(40)
        self.output_btn.setStyleSheet("""
            QPushButton {
                font-size: 18px;
                font-weight: bold;
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        self.output_btn.clicked.connect(self.select_output_path)
        file_btn_layout.addWidget(self.input_btn)
        file_btn_layout.addWidget(self.output_btn)
        file_layout.addLayout(file_btn_layout)
        # 文件信息显示
        self.file_info_label = QLabel("� 未选择文件")
        self.file_info_label.setStyleSheet("color: #1E3A8A; font-size: 10pt; padding: 5px; background-color: #f8f9fa; border-radius: 3px;")
        file_layout.addWidget(self.file_info_label)
        # 内容交互区
        self.content_tabs = QTabWidget()
        self.content_tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 2px solid #bdc3c7;
                border-radius: 5px;
                background-color: white;
            }
            QTabBar::tab {
                background-color: #ecf0f1;
                padding: 8px 16px;
                margin-right: 2px;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
            }
            QTabBar::tab:selected {
                background-color: #3498db;
                color: white;
            }
        """)

        self.text_tab = QWidget()
        self.excel_tab = QWidget()
        self.word_tab = QWidget()

        # 文本内容交互
        text_layout = QVBoxLayout()
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(False)  # 允许编辑以支持交互式脱敏
        # QTextEdit用独立的预览提示框代替setPlaceholderText，保持与Excel风格统一
        text_placeholder = QLabel("📄 选择文本文件后，内容将在此显示\n💡 使用技巧：\n• 选中需要脱敏的文字后右键选择脱敏方式\n• 支持局部脱敏和全文同内容脱敏")
        text_placeholder.setAlignment(Qt.AlignmentFlag.AlignLeft)
        text_placeholder.setStyleSheet("color: #2563EB; font-size: 9pt; background-color: #f8f9fa; border: 1px solid #e9ecef; border-radius: 4px; padding: 10px;")
        text_layout.addWidget(text_placeholder)

        # 添加文本选择上下文菜单
        self.text_edit.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.text_edit.customContextMenuRequested.connect(self.show_text_context_menu)

        text_layout.addWidget(self.text_edit)
        
        # 初始化文本右键菜单
        self.text_menu = QMenu(self)
        self.redact_action = QAction("🎯 标记脱敏（仅选中部分）", self)
        self.redact_action.triggered.connect(self.mark_text_redaction)
        self.text_menu.addAction(self.redact_action)
        
        self.redact_all_action = QAction("🔄 标记脱敏（全文相同内容）", self)
        self.redact_all_action.triggered.connect(self.mark_text_redaction_all)
        self.text_menu.addAction(self.redact_all_action)
        
        # 添加撤销脱敏功能
        self.text_menu.addSeparator()
        self.text_undo_action = QAction("↩️ 撤销脱敏", self)
        self.text_undo_action.triggered.connect(self.undo_text_redaction)
        self.text_menu.addAction(self.text_undo_action)
        
        # 区域撤销功能已移除，仅保留单步撤销
        
        # Excel内容交互
        excel_layout = QVBoxLayout()
        self.table_widget = QTableWidget()
        # QTableWidget没有setPlaceholderText方法，我们用标签代替
        excel_placeholder = QLabel("📊 选择Excel文件后，内容将在此显示\n💡 使用技巧：\n• 点击单元格选中后右键选择脱敏方式\n• 支持单元格、整行、整列脱敏 ")
        excel_placeholder.setAlignment(Qt.AlignmentFlag.AlignLeft)
        excel_placeholder.setStyleSheet("color: #2563EB; font-size: 9pt; background-color: #f8f9fa; border: 1px solid #e9ecef; border-radius: 4px; padding: 10px;")
        excel_layout.addWidget(excel_placeholder)
        excel_layout.addWidget(self.table_widget)
        
        # Word文档内容交互
        word_layout = QVBoxLayout()
        self.word_edit = QTextEdit()
        self.word_edit.setReadOnly(False)  # 允许编辑以支持交互式脱敏
        # QTextEdit用独立的预览提示框代替setPlaceholderText，保持与Excel风格统一
        word_placeholder = QLabel("📝 选择Word文档后，内容将在此显示\n💡 使用技巧：\n• 选中需要脱敏的文字后右键选择脱敏方式\n• 支持局部脱敏和全文同内容脱敏 ")
        word_placeholder.setAlignment(Qt.AlignmentFlag.AlignLeft)
        word_placeholder.setStyleSheet("color: #2563EB; font-size: 9pt; background-color: #f8f9fa; border: 1px solid #e9ecef; border-radius: 4px; padding: 10px;")
        word_layout.addWidget(word_placeholder)
        
        # 添加Word文档选择上下文菜单
        self.word_edit.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.word_edit.customContextMenuRequested.connect(self.show_word_context_menu)
        
        word_layout.addWidget(self.word_edit)
        
        # 初始化Word右键菜单
        self.word_menu = QMenu(self)
        self.word_redact_action = QAction("🎯 标记脱敏（仅选中部分）", self)
        self.word_redact_action.triggered.connect(self.mark_word_redaction)
        self.word_menu.addAction(self.word_redact_action)
        
        self.word_redact_all_action = QAction("🔄 标记脱敏（全文相同内容）", self)
        self.word_redact_all_action.triggered.connect(self.mark_word_redaction_all)
        self.word_menu.addAction(self.word_redact_all_action)
        
        # 添加撤销脱敏功能
        self.word_menu.addSeparator()
        self.word_undo_action = QAction("↩️ 撤销脱敏", self)
        self.word_undo_action.triggered.connect(self.undo_word_redaction)
        self.word_menu.addAction(self.word_undo_action)
        
        # 区域撤销功能已移除，仅保留单步撤销
        
        self.text_tab.setLayout(text_layout)
        self.excel_tab.setLayout(excel_layout)
        self.word_tab.setLayout(word_layout)
        self.content_tabs.addTab(self.word_tab, "📝 Word文档")
        self.content_tabs.addTab(self.excel_tab, "📊 Excel内容")
        self.content_tabs.addTab(self.text_tab, "📄 文本内容")
        file_layout.addWidget(self.content_tabs)
        
        # 初始化表格右键菜单
        self.setup_table_context_menu()
        
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)

        # 操作按钮区
        action_btn_layout = QHBoxLayout()
        self.process_btn = QPushButton("🚀 开始脱敏")
        self.process_btn.setMinimumHeight(50)
        self.process_btn.setStyleSheet("""
            QPushButton {
                font-size: 20px;
                font-weight: bold;
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 15px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        
        self.batch_btn = QPushButton("📦 批量处理")
        self.batch_btn.setMinimumHeight(50)
        self.batch_btn.setStyleSheet("""
            QPushButton {
                font-size: 20px;
                font-weight: bold;
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 15px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        self.batch_btn.setVisible(False)  # 初始隐藏，只在自动规则模式下显示
        
        self.help_btn = QPushButton("❓ 帮助")
        self.help_btn.setMinimumHeight(50)
        self.help_btn.setStyleSheet("""
            QPushButton {
                font-size: 20px;
                font-weight: bold;
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 15px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        
        action_btn_layout.addWidget(self.process_btn)
        action_btn_layout.addWidget(self.batch_btn)
        action_btn_layout.addWidget(self.help_btn)
        main_layout.addLayout(action_btn_layout)
        
        # 连接所有按钮
        self.process_btn.clicked.connect(self.process_file)
        self.batch_btn.clicked.connect(self.batch_process)
        self.help_btn.clicked.connect(self.show_help)

        # 进度条和状态栏
        progress_layout = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #bdc3c7;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                border-radius: 3px;
            }
        """)
        self.status_label = QLabel("✅ 就绪")
        self.status_label.setStyleSheet("color: #27ae60; font-weight: bold;")
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_label)
        main_layout.addLayout(progress_layout)

        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
        # 初始化规则引擎（隐藏在后台）
        self.rule_engine = RuleEngine()
        self.rule_engine.load_default_rules()
        
        # 验证规则示例（仅在开发调试时使用）
        # self.rule_engine.verify_rule_examples()

    def show_help(self):
        # 更新记录（简洁版，每条一句话，保留日期）
        update_records = [
            "2025-08-11 V2.1.5：增加Excel区域撤销功能，提升一致性和稳定性。",
            "2025-08-11 V2.1.4：Excel批量脱敏和全表查找替换功能增强。",
            "2025-08-11 V2.1.3：新增检查更新功能，支持自动获取最新版本",
            "2025-08-11 V2.1.2：新增自定义字段规则，界面优化",
            "2025-08-11 V2.1.1：修复姓名脱敏逻辑，批量粘贴更智能",
            "2025-08-10 V2.1.0：规则选择弹窗升级，界面统一化",
            "2025-08-09 V2.0.0：界面重构，支持Word/Excel批量处理",
            "2025-08-08 V1.9.3：Excel交互增强，视觉优化",
            "2025-08-04 V1.0：初始版本，基础UI与文件操作"
        ]
        help_text = f"""
        <div style='max-width:900px; margin:auto; font-family:Arial;'>
        <div style='color: #c0392b; font-weight: bold; border: 10px solid #c0392b; padding: 20px; margin-bottom: 25px; font-size:20px;'>
        【免责声明】本软件为免费工具，用户自愿使用。开发者不承诺软件绝对安全，对因使用软件导致的数据丢失、系统损坏等后果不承担责任。禁止将软件用于非法目的。
        </div>

    <h2 style='text-align:center; font-size:28px; margin-bottom:18px;'>通用脱敏工具 V2.1.5 使用说明</h2>

        <h3 style='color:#2980b9; font-size:25px;'>基本功能</h3>
        <ul style='font-size:18px;'>
            <li>支持 TXT文本、Excel表格、Word文档三种格式的敏感信息脱敏处理</li>
            <li>交互式脱敏：选中文本或单元格，右键标记，精确控制每个内容</li>
            <li>自动脱敏（规则模式）：配置规则后可一键批量处理文件夹或多文件</li>
        </ul>

        <h3 style='color:#2980b9; font-size:25px;'>核心特色</h3>
        <ul style='font-size:18px;'>
            <li>右键快速标记，支持全文同步脱敏</li>
            <li>Excel支持单元格、整行、整列精确脱敏</li>
            <li>自定义规则每日自动保存，支持批量导入/导出</li>
            <li>内置13种脱敏规则，涵盖生活工作多方面需求</li>
        </ul>

        <h3 style='color:#2980b9; font-size:25px;'>操作步骤</h3>
        <ol style='font-size:18px;'>
            <li>选择文件或文件夹：支持单文件、多文件或文件夹批量处理</li>
            <li>设置输出路径：可自定义输出目录，自动生成“文件名（脱敏）”格式</li>
            <li>配置脱敏规则：可自定义姓名、字段等规则，支持实时预览</li>
            <li>开始处理：点击“开始脱敏”或“批量处理”按钮自动完成脱敏</li>
        </ol>

        <div style='display:flex; align-items:center;'>
            <h3 style='color:#2980b9; font-size:25px; margin:0;'>更新记录</h3>
        </div>
        <div style='max-height:160px; overflow-y:auto; background:#f8f9fa; border:1px solid #e9ecef; border-radius:6px; padding:8px; font-size:18px;'>
        <ul style='margin:0;'>
            {''.join([f'<li>{rec}</li>' for rec in update_records])}
        </ul>
        </div>

        <div style='background-color:#f0f8ff; padding:10px; margin:14px 0; border-left:4px solid #4a90e2; font-size:20px;'>
        <b>使用建议</b><br>
        • 新手建议先用交互式模式熟悉功能<br>
        • 批量处理前建议先单个文件测试<br>
        • 重要文件处理前请务必备份
        </div>

        <p style='text-align:center; margin-top:18px; font-size:20px;'>
        <b>版本 V2.1.5</b> | 2025 D&Ai <br>
        <b>更多功能请在使用中探索发现 😊</b>
        </p>
        </div>
        """

        # 使用 QDialog + QScrollArea + QLabel 实现带滚动条的帮助窗口
        help_window = QDialog(self)
        help_window.setWindowTitle("帮助")
        help_window.setWindowModality(Qt.WindowModality.ApplicationModal)
        help_window.resize(900, 600)

        # 主布局
        main_layout = QVBoxLayout(help_window)
        
        # 顶部按钮区域
        top_layout = QHBoxLayout()
        top_layout.addStretch()  # 左侧弹簧
        
        check_update_btn = QPushButton("检查更新")
        check_update_btn.setStyleSheet("""
            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                padding: 8px 20px;
                font-size: 14px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #357abd;
            }
            QPushButton:pressed {
                background-color: #2968a3;
            }
        """)
        check_update_btn.clicked.connect(self.check_update)
        top_layout.addWidget(check_update_btn)
        
        main_layout.addLayout(top_layout)

        scroll_area = QScrollArea(help_window)
        scroll_area.setWidgetResizable(True)

        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)

        help_label = QLabel(help_text)
        help_label.setWordWrap(True)
        help_label.setStyleSheet("font-size: 14px; padding: 10px;")
        help_label.setTextFormat(Qt.TextFormat.RichText)
        content_layout.addWidget(help_label)

        scroll_area.setWidget(content_widget)

        main_layout.addWidget(scroll_area)
        help_window.setLayout(main_layout)

        help_window.exec_()

    def update_rule_list(self):
        """更新规则总览（弹窗显示，避免主窗口属性缺失）"""
        total_rules = len(self.rule_engine.rules)
        active_rules_list = self.rule_engine.get_active_rules()
        active_count = len(active_rules_list)
        if active_count == 0:
            msg = f"当前未激活任何规则，共 {total_rules} 条"
        else:
            names = '\n'.join([f"{i+1}. {r.name}" for i, r in enumerate(active_rules_list)])
            msg = f"已激活 {active_count} 条规则，共 {total_rules} 条：\n\n{names}"
        QMessageBox.information(self, "规则设置", msg)

    def add_rule(self):
        """添加新规则"""
        rule_text = self.rule_edit.toPlainText().strip()
        if not rule_text:
            QMessageBox.warning(self, "警告", "请输入规则内容")
            return
        
        try:
            # 简单示例：实际应解析JSON格式规则
            new_rule = RedactionRule(
                rule_id=f"rule_{len(self.rule_engine.rules)+1}",
                name="自定义规则",
                pattern=rule_text,
                example=f"示例: 应用 {rule_text}"
            )
            self.rule_engine.add_rule(new_rule)
            self.update_rule_list()
            self.rule_edit.clear()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"添加规则失败: {str(e)}")

    def clear_rules(self):
        """清空所有规则"""
        reply = QMessageBox.question(
            self, 
            "确认", 
            "确定要清空所有规则吗?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.rule_engine.rules.clear()
            self.update_rule_list()

    def export_rules(self):
        """导出规则到JSON文件"""
        if not self.rule_engine.rules:
            QMessageBox.warning(self, "警告", "没有可导出的规则")
            return
        
        # 默认保存到 user_custom_rules 文件夹
        import os, datetime
        rules_dir = os.path.join(os.path.dirname(__file__), "user_custom_rules")
        if not os.path.exists(rules_dir):
            os.makedirs(rules_dir)
        
        date_str = datetime.datetime.now().strftime("%Y%m%d")
        default_filename = os.path.join(rules_dir, f"导出规则{date_str}.json")
            
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "导出规则",
            default_filename,
            "JSON文件 (*.json);;所有文件 (*)"
        )
        
        if file_path:
            try:
                import json
                rules_data = [{
                    'rule_id': rule.rule_id,
                    'name': rule.name,
                    'pattern': rule.pattern,
                    'example': rule.example,
                    'is_active': rule.is_active
                } for rule in self.rule_engine.rules]
                
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(rules_data, f, ensure_ascii=False, indent=2)
                    
                QMessageBox.information(self, "成功", "规则导出成功")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")

    def import_rules(self):
        """从JSON文件导入规则"""
        # 默认打开 user_custom_rules 文件夹
        import os
        rules_dir = os.path.join(os.path.dirname(__file__), "user_custom_rules")
        if not os.path.exists(rules_dir):
            os.makedirs(rules_dir)
        
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "导入规则",
            rules_dir,
            "JSON文件 (*.json);;所有文件 (*)"
        )
        
        if file_path:
            try:
                import json
                with open(file_path, 'r', encoding='utf-8') as f:
                    rules_data = json.load(f)
                
                # 检查是否是数组格式
                if not isinstance(rules_data, list):
                    # 如果是对象格式，检查是否有rules字段
                    if isinstance(rules_data, dict) and 'rules' in rules_data:
                        rules_data = rules_data['rules']
                    else:
                        raise ValueError("JSON格式不正确，应为规则数组或包含rules字段的对象")
                
                # 询问用户是否要替换现有规则还是添加到现有规则
                reply = QMessageBox.question(
                    self, 
                    "导入选项", 
                    f"检测到 {len(rules_data)} 条规则\n\n是否要替换现有的内置规则？\n\n点击 '是' 替换所有规则\n点击 '否' 添加到现有规则中\n点击 '取消' 取消导入",
                    QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
                )
                
                if reply == QMessageBox.Cancel:
                    return
                
                new_rules = []
                for rule_data in rules_data:
                    new_rule = RedactionRule(
                        rule_id=rule_data.get('id', rule_data.get('rule_id', f"imported_{len(new_rules)+1}")),
                        name=rule_data.get('name', '导入规则'),
                        pattern=rule_data.get('pattern', ''),
                        example=rule_data.get('example', ''),
                        regex=rule_data.get('pattern', ''),
                        marker_char="*",
                        is_active=rule_data.get('is_active', True)
                    )
                    new_rules.append(new_rule)
                
                if reply == QMessageBox.Yes:
                    # 替换现有规则
                    self.rule_engine.rules = new_rules
                    action = "替换"
                else:
                    # 添加到现有规则
                    self.rule_engine.rules.extend(new_rules)
                    action = "添加"
                
                self.update_rule_list()
                active_count = len([r for r in new_rules if r.is_active])
                QMessageBox.information(
                    self, 
                    "成功", 
                    f"成功{action} {len(new_rules)} 条规则\n其中 {active_count} 条已激活\n当前总规则数: {len(self.rule_engine.rules)}"
                )
                
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导入失败: {str(e)}")

    def edit_rule(self):
        """编辑现有规则"""
        if not self.rule_engine.rules:
            QMessageBox.warning(self, "警告", "没有可编辑的规则")
            return
            
        # 创建规则选择对话框
        rule_names = [f"{i+1}. {rule.name} ({'✅' if rule.is_active else '❌'})" for i, rule in enumerate(self.rule_engine.rules)]
        rule_name, ok = QInputDialog.getItem(
            self,
            "选择规则",
            "请选择要编辑的规则:",
            rule_names,
            0,
            False
        )
        
        if ok and rule_name:
            try:
                # 获取选中的规则索引
                rule_index = int(rule_name.split('.')[0]) - 1
                selected_rule = self.rule_engine.rules[rule_index]
                
                # 创建编辑对话框
                dialog = QDialog(self)
                dialog.setWindowTitle("编辑规则")
                dialog.setMinimumWidth(500)
                layout = QVBoxLayout()
                
                # 规则名称
                name_label = QLabel("规则名称:")
                name_edit = QLineEdit(selected_rule.name)
                layout.addWidget(name_label)
                layout.addWidget(name_edit)
                
                # 规则模式
                pattern_label = QLabel("匹配模式:")
                pattern_edit = QLineEdit(selected_rule.pattern)
                layout.addWidget(pattern_label)
                layout.addWidget(pattern_edit)
                
                # 示例
                example_label = QLabel("示例:")
                example_edit = QLineEdit(selected_rule.example)
                layout.addWidget(example_label)
                layout.addWidget(example_edit)
                
                # 规则状态区域
                status_layout = QHBoxLayout()
                status_label = QLabel("规则状态:")
                status_layout.addWidget(status_label)
                
                # 激活/禁用按钮
                toggle_btn = QPushButton()
                if selected_rule.is_active:
                    toggle_btn.setText("🔴 禁用规则")
                    toggle_btn.setStyleSheet("background-color: #ff6b6b; color: white; font-weight: bold;")
                else:
                    toggle_btn.setText("🟢 激活规则")
                    toggle_btn.setStyleSheet("background-color: #51cf66; color: white; font-weight: bold;")
                
                def toggle_rule_status():
                    selected_rule.is_active = not selected_rule.is_active
                    if selected_rule.is_active:
                        toggle_btn.setText("🔴 禁用规则")
                        toggle_btn.setStyleSheet("background-color: #ff6b6b; color: white; font-weight: bold;")
                    else:
                        toggle_btn.setText("🟢 激活规则")
                        toggle_btn.setStyleSheet("background-color: #51cf66; color: white; font-weight: bold;")
                
                toggle_btn.clicked.connect(toggle_rule_status)
                status_layout.addWidget(toggle_btn)
                status_layout.addStretch()
                
                layout.addLayout(status_layout)
                
                # 按钮
                btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
                btn_box.accepted.connect(dialog.accept)
                btn_box.rejected.connect(dialog.reject)
                layout.addWidget(btn_box)
                
                dialog.setLayout(layout)
                
                if dialog.exec_() == QDialog.Accepted:
                    # 更新规则
                    selected_rule.name = name_edit.text()
                    selected_rule.pattern = pattern_edit.text()
                    selected_rule.regex = pattern_edit.text()  # 同时更新regex
                    selected_rule.example = example_edit.text()
                    self.update_rule_list()
                    QMessageBox.information(self, "成功", "规则更新成功")
                    
            except Exception as e:
                QMessageBox.critical(self, "错误", f"编辑规则失败: {str(e)}")

    def preview_rule(self):
        """预览规则效果"""
        if not self.rule_engine.rules:
            QMessageBox.warning(self, "警告", "没有可预览的规则")
            return
            
        # 创建预览对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("规则预览")
        dialog.resize(500, 400)
        layout = QVBoxLayout()
        
        # 规则选择下拉框
        rule_combo = QComboBox()
        rule_combo.addItems([rule.name for rule in self.rule_engine.rules])
        layout.addWidget(QLabel("选择规则:"))
        layout.addWidget(rule_combo)
        
        # 测试文本输入
        test_text_edit = QTextEdit()
        test_text_edit.setPlaceholderText("输入测试文本...")
        layout.addWidget(QLabel("测试文本:"))
        layout.addWidget(test_text_edit)
        
        # 预览结果
        result_label = QLabel("预览结果:")
        result_text = QTextEdit()
        result_text.setReadOnly(True)
        layout.addWidget(result_label)
        layout.addWidget(result_text)
        
        # 预览按钮
        preview_btn = QPushButton("预览")
        def on_preview():
            selected_index = rule_combo.currentIndex()
            if selected_index >= 0:
                rule = self.rule_engine.rules[selected_index]
                test_text = test_text_edit.toPlainText()
                if not test_text:
                    result_text.setPlainText("请输入测试文本")
                    return
                    
                # 使用与交互脱敏相同的智能脱敏逻辑
                try:
                    import re
                    # 使用规则引擎的模式匹配，但结合智能脱敏逻辑
                    # 使用 re.finditer 来获取完整匹配，避免分组问题
                    matches = []
                    for match_obj in re.finditer(rule.pattern, test_text):
                        matches.append(match_obj.group(0))  # group(0) 是完整匹配
                    
                    if not matches:
                        result = test_text
                        matches_found = False
                    else:
                        result = test_text
                        matches_found = True
                        for match in matches:
                            # 使用smart_redact_text的智能脱敏逻辑
                            redacted = self.smart_redact_text(match)
                            result = result.replace(match, redacted)
                    
                    # 显示详细结果
                    status = "✅ 匹配成功" if matches_found else "❌ 未匹配"
                    match_info = f"匹配到的内容: {matches}" if matches_found else "无匹配内容"
                    result_text.setPlainText(f"""应用规则: {rule.name}
规则描述: {rule.example}
匹配状态: {status}
{match_info}

原始文本: {test_text}

脱敏结果: {result}

注：此预览使用与实际脱敏相同的算法""")
                    
                except Exception as e:
                    result_text.setPlainText(f"预览出错: {str(e)}\n\n请检查规则格式是否正确")
                    
        preview_btn.clicked.connect(on_preview)
        layout.addWidget(preview_btn)
        
        dialog.setLayout(layout)
        dialog.exec_()

    def read_file_with_encoding(self, file_path):
        """尝试多种编码格式读取文件"""
        encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig', 'latin1']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    content = f.read()
                    self.original_encoding = encoding  # 保存原始编码
                    self.status_label.setText(f"使用 {encoding} 编码成功读取文件")
                    return content
            except UnicodeDecodeError:
                continue
            except Exception as e:
                self.status_label.setText(f"读取文件时发生错误: {str(e)}")
                continue
        
        return None

    def load_word_document(self, file_path):
        """加载Word文档内容"""
        try:
            if file_path.lower().endswith('.docx'):
                try:
                    from docx import Document
                    doc = Document(file_path)
                    full_text = []
                    for para in doc.paragraphs:
                        full_text.append(para.text)
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                full_text.append(cell.text)
                    self.current_word_doc = doc
                    return '\n'.join(full_text)
                except ImportError:
                    QMessageBox.warning(self, "警告", "未安装python-docx库，无法处理DOCX文件\n请运行: pip install python-docx")
                    return None
            else:
                QMessageBox.warning(self, "警告", "仅支持DOCX格式文件，请先转换DOC为DOCX后再处理。")
                return None
        except Exception as e:
            QMessageBox.warning(self, "警告", f"加载Word文档失败: {str(e)}")
            return None

    def on_mode_changed(self):
        """处理模式切换"""
        if self.mode_combo.currentIndex() == 0:
            # 交互式脱敏模式
            self.mode_tip_label.setText("💡 交互式脱敏：选中文本或单元格后右键选择脱敏，精确控制每个内容")
            self.mode_tip_label.setStyleSheet("color: #27ae60; font-size: 9pt; background-color: #d5f4e6; padding: 8px; border-radius: 5px; border-left: 3px solid #27ae60;")
            self.rule_config_btn.setVisible(False)  # 隐藏规则配置按钮
            self.batch_btn.setVisible(False)  # 隐藏批量处理按钮
        else:
            # 自动脱敏模式
            self.mode_tip_label.setText("⚙️ 自动脱敏：将对整个文件应用脱敏规则，请点击【配置脱敏规则】按钮设置规则")
            self.mode_tip_label.setStyleSheet("color: #e74c3c; font-size: 9pt; background-color: #fdf2f2; padding: 8px; border-radius: 5px; border-left: 3px solid #e74c3c;")
            self.rule_config_btn.setVisible(True)  # 显示规则配置按钮
            self.batch_btn.setVisible(True)  # 显示批量处理按钮

    def show_rule_config_dialog(self):
        """显示规则配置弹窗（菜单式复选框，每条规则可勾选激活/禁用）"""
        dialog = QDialog(self)
        dialog.setWindowTitle("📋 脱敏规则配置（可视化选择）")
        dialog.setModal(True)
        dialog.resize(700, 600)

        layout = QVBoxLayout()

        # 规则复选框列表区
        rules_group = QGroupBox("内置规则选择（勾选表示激活）")
        rules_layout = QVBoxLayout()
        self.rule_checkboxes = []
        for i, rule in enumerate(self.rule_engine.rules):
            cb = QCheckBox(f"{i+1}. {rule.name}  ——  {rule.example}")
            cb.setChecked(rule.is_active)
            cb.setToolTip(f"匹配规则: {rule.pattern}")
            self.rule_checkboxes.append(cb)
            row_layout = QHBoxLayout()
            row_layout.addWidget(cb)
            # 如果是姓名规则，在后面加自定义按钮
            if rule.name == "姓名":
                name_btn = QPushButton("自定义名单")
                name_btn.setStyleSheet("""
                    QPushButton {
                        font-size: 12px;
                        font-weight: bold;
                        background-color: #3498db;
                        color: white;
                        border: none;
                        border-radius: 5px;
                        padding: 6px 12px;
                        margin-left: 10px;
                    }
                    QPushButton:hover {
                        background-color: #2980b9;
                    }
                """)
                name_btn.setCursor(Qt.CursorShape.PointingHandCursor)
                name_btn.clicked.connect(self.show_name_redact_dialog)
                row_layout.addWidget(name_btn)
            # 如果是自定义字段规则，在后面加自定义按钮
            elif rule.name == "自定义字段":
                field_btn = QPushButton("自定义字段")
                field_btn.setStyleSheet("""
                    QPushButton {
                        font-size: 12px;
                        font-weight: bold;
                        background-color: #e67e22;
                        color: white;
                        border: none;
                        border-radius: 5px;
                        padding: 6px 12px;
                        margin-left: 10px;
                    }
                    QPushButton:hover {
                        background-color: #d35400;
                    }
                """)
                field_btn.setCursor(Qt.CursorShape.PointingHandCursor)
                field_btn.clicked.connect(self.show_custom_field_redact_dialog)
                row_layout.addWidget(field_btn)
            rules_layout.addLayout(row_layout)
        rules_group.setLayout(rules_layout)
        layout.addWidget(rules_group)

        # 规则编辑区（保留原有功能）
        self.rule_edit = QTextEdit()
        placeholder_text = """请输入脱敏规则(JSON格式)...\n\n📝 使用步骤提示：\n第1步：确定需要脱敏的敏感信息类型（如：姓名、电话、身份证等）\n第2步：为每个类型编写匹配规则（支持正则表达式）\n第3步：设置替换方式（如：张三 → 张XX，13812345678 → 138****5678）\n\n💡 示例格式：\n[\n    {\n        \"name\": \"姓名脱敏\",\n        \"pattern\": \"张三|李四|王五\",\n        \"replacement\": \"***\",\n        \"is_regex\": false\n    },\n    {\n        \"name\": \"手机号脱敏\", \n        \"pattern\": \"1[3-9]\\\\d{9}\",\n        \"replacement\": \"***\",\n        \"is_regex\": true\n    }\n]\n\n💭 小贴士：可以使用下方按钮导入已有规则文件或使用预览功能测试效果"""
        self.rule_edit.setPlaceholderText(placeholder_text)
        layout.addWidget(QLabel("高级规则编辑区:"))
        layout.addWidget(self.rule_edit)

        # 规则操作按钮（保留原有功能）
        rule_btn_layout = QHBoxLayout()
        add_btn = QPushButton("➕ 添加规则")
        add_btn.clicked.connect(self.add_rule)
        import_btn = QPushButton("📥 导入规则")
        import_btn.clicked.connect(self.import_rules)
        export_btn = QPushButton("📤 导出规则")
        export_btn.clicked.connect(self.export_rules)
        edit_btn = QPushButton("✏️ 编辑规则")
        edit_btn.clicked.connect(self.edit_rule)
        preview_btn = QPushButton("👁️ 预览规则")
        preview_btn.clicked.connect(self.preview_rule)
        clear_btn = QPushButton("🗑️ 清空规则")
        clear_btn.clicked.connect(self.clear_rules)
        rule_btn_layout.addWidget(add_btn)
        rule_btn_layout.addWidget(import_btn)
        rule_btn_layout.addWidget(export_btn)
        rule_btn_layout.addWidget(edit_btn)
        rule_btn_layout.addWidget(preview_btn)
        rule_btn_layout.addWidget(clear_btn)
        layout.addLayout(rule_btn_layout)

        # 对话框按钮
        dialog_btn_layout = QHBoxLayout()
        dialog_btn_layout.addStretch()
        
        ok_btn = QPushButton("✅ 继续")
        ok_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 15px 30px;
                font-size: 18px;
                font-weight: bold;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #229954;
            }
        """)
        
        cancel_btn = QPushButton("❌ 取消")
        cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 15px 30px;
                font-size: 18px;
                font-weight: bold;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        
        dialog_btn_layout.addWidget(ok_btn)
        dialog_btn_layout.addWidget(cancel_btn)
        dialog_btn_layout.addStretch()
        layout.addLayout(dialog_btn_layout)

        # 按钮事件：保存激活状态
        def on_ok():
            for cb, rule in zip(self.rule_checkboxes, self.rule_engine.rules):
                rule.is_active = cb.isChecked()
            active_count = len([r for r in self.rule_engine.rules if r.is_active])
            QMessageBox.information(self, "规则设置", f"已激活 {active_count} 条规则")
            self.update_rule_list()
            dialog.accept()
        ok_btn.clicked.connect(on_ok)
        cancel_btn.clicked.connect(dialog.reject)

        dialog.setLayout(layout)
        dialog.exec_()

    def sanitize_excel_value(self, value):
        """清理Excel单元格值，避免特殊字符问题"""
        if value is None:
            return ""
        
        # 转换为字符串
        str_value = str(value)
        
        # 替换可能导致问题的字符
        # Excel不允许某些字符，特别是控制字符
        forbidden_chars = ['\x00', '\x01', '\x02', '\x03', '\x04', '\x05', '\x06', '\x07', '\x08']
        for char in forbidden_chars:
            str_value = str_value.replace(char, '')
        
        # 限制长度（Excel单元格最大32767字符）
        if len(str_value) > 32767:
            str_value = str_value[:32767]
        
        return str_value

    def select_input_file(self):
        """选择输入文件并加载内容"""
        # 根据当前标签页确定文件类型过滤器
        current_tab = self.content_tabs.currentIndex()
        if current_tab == 0:  # Word标签页
            file_filter = "Word文档 (*.docx *.doc);;Word 2007及以上 (*.docx);;Word 97-2003 (*.doc);;所有文件 (*)"
            dialog_title = "选择Word文档"
        elif current_tab == 1:  # Excel标签页
            file_filter = "Excel文件 (*.xlsx);;所有文件 (*)"
            dialog_title = "选择Excel文件"
        elif current_tab == 2:  # 文本标签页
            file_filter = "文本文件 (*.txt);;所有文件 (*)"
            dialog_title = "选择文本文件"
        else:
            file_filter = "Word文档 (*.docx *.doc);;Excel文件 (*.xlsx);;文本文件 (*.txt);;所有文件 (*)"
            dialog_title = "选择输入文件"
            
        script_dir = os.path.dirname(os.path.abspath(__file__))
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            dialog_title,
            script_dir,
            file_filter
        )
        if file_path:
            self.input_file_path = file_path
            self.file_info_label.setText(f"输入文件: {os.path.basename(file_path)}")
            self.status_label.setText("已选择输入文件")
            
            # 根据文件类型加载内容
            try:
                if file_path.lower().endswith('.txt'):
                    # 尝试多种编码格式
                    content = self.read_file_with_encoding(file_path)
                    if content is not None:
                        self.text_edit.setPlainText(content)
                        self.content_tabs.setCurrentIndex(2)  # 切换到文本标签页
                    else:
                        QMessageBox.warning(self, "警告", "无法读取文件，请检查文件编码格式")
                        return
                elif file_path.lower().endswith('.xlsx'):
                    try:
                        from openpyxl import load_workbook
                        
                        wb = load_workbook(file_path)
                        if wb and wb.active:
                            # 保存工作簿信息以便后续保存
                            self.current_workbook = wb
                            self.current_sheet_name = wb.active.title
                            self.original_excel_path = file_path  # 保存原始文件路径
                            
                            ws = wb.active
                            
                            # 清空表格和格式存储
                            self.table_widget.clear()
                            self.excel_cell_formats = {}  # 清空格式存储
                            
                            # 获取行数和列数
                            max_row = ws.max_row if hasattr(ws, 'max_row') else 0
                            max_col = ws.max_column if hasattr(ws, 'max_column') else 0
                            self.table_widget.setRowCount(max_row)
                            self.table_widget.setColumnCount(max_col)
                            
                            # 加载Excel数据并保存格式信息
                            try:
                                for row_idx, row in enumerate(ws.iter_rows(), 1):
                                    for col_idx, cell in enumerate(row, 1):
                                        if cell and cell.value is not None:
                                            try:
                                                value = str(cell.value) if cell.value is not None else ""
                                                self.table_widget.setItem(
                                                    row_idx - 1,  # QTableWidget从0开始
                                                    col_idx - 1,
                                                    QTableWidgetItem(value))
                                                
                                                # 保存单元格格式信息
                                                self.save_cell_format(cell, row_idx - 1, col_idx - 1)
                                            except Exception:
                                                continue
                            except Exception as e:
                                QMessageBox.warning(self, "警告", f"加载Excel数据失败: {str(e)}")
                                return
                            
                            self.content_tabs.setCurrentIndex(1)
                        else:
                            QMessageBox.warning(self, "警告", "Excel文件中没有活动工作表")
                    except Exception as e:
                        QMessageBox.warning(self, "警告", f"加载Excel文件失败: {str(e)}")
                elif file_path.lower().endswith(('.docx', '.doc')):
                    try:
                        # .doc 文件仅提示用户先转换为 .docx，避免直接处理
                        if file_path.lower().endswith('.doc'):
                            QMessageBox.information(
                                self,
                                "格式提示",
                                "很抱歉，由于兼容性原因，暂不支持直接处理DOC格式文件。\n请先使用word将DOC文件另存为DOCX格式后再继续脱敏工作。"
                            )
                            return

                        # 处理DOCX文档
                        content = self.load_word_document(file_path)
                        if content is not None:
                            self.word_edit.setPlainText(content)
                            self.content_tabs.setCurrentIndex(0)  # 切换到Word标签页
                        else:
                            QMessageBox.warning(self, "警告", "无法读取DOCX文档")
                            return
                    except Exception as e:
                        QMessageBox.warning(self, "警告", f"加载Word文档失败: {str(e)}")
                else:
                    QMessageBox.warning(self, "警告", "不支持的文件格式")
            except Exception as e:
                QMessageBox.warning(self, "警告", f"加载文件失败: {str(e)}")

    def select_output_path(self):
        """设置输出路径"""
        if not hasattr(self, 'input_file_path'):
            QMessageBox.warning(self, "警告", "请先选择输入文件")
            return
            
        # 获取输入文件信息
        input_dir = os.path.dirname(self.input_file_path)
        input_name = os.path.splitext(os.path.basename(self.input_file_path))[0]
        input_ext = os.path.splitext(self.input_file_path)[1]
        
        # 生成默认输出文件名
        default_name = f"{input_name}（脱敏）{input_ext}"
        default_path = os.path.join(input_dir, default_name)
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "设置输出路径",
            default_path,  # 设置默认路径和文件名
            f"相同类型文件 (*{input_ext});;所有文件 (*)"
        )
        if file_path:
            # 确保输出文件扩展名与输入文件一致
            output_ext = os.path.splitext(file_path)[1]
            if output_ext.lower() != input_ext.lower():
                file_path = file_path + input_ext
            self.output_file_path = file_path
            self.file_info_label.setText(f"{self.file_info_label.text()}\n输出路径: {os.path.basename(file_path)}")
            self.status_label.setText("已设置输出路径")

    def process_file(self):
        """执行文件脱敏处理"""
        if not hasattr(self, 'input_file_path'):
            QMessageBox.warning(self, "警告", "请先选择输入文件")
            return
            
        if not hasattr(self, 'output_file_path'):
            QMessageBox.warning(self, "警告", "请先设置输出路径")
            return
        
        # 检查处理模式
        is_interactive_mode = self.mode_combo.currentIndex() == 0
        
        if is_interactive_mode:
            # 交互式脱敏模式 - 只保存当前界面上的修改
            self.save_interactive_changes()
        else:
            # 自动脱敏模式 - 应用规则到整个文件
            if not self.rule_engine.get_active_rules():
                QMessageBox.warning(self, "警告", "没有可用的脱敏规则")
                return
            self.auto_process_file()
    
    def save_interactive_changes(self):
        """保存交互式脱敏的更改"""
        try:
            # 如果没有设置输出路径，自动生成默认路径
            if not hasattr(self, 'output_file_path') or not self.output_file_path:
                input_dir = os.path.dirname(self.input_file_path)
                input_name = os.path.splitext(os.path.basename(self.input_file_path))[0]
                input_ext = os.path.splitext(self.input_file_path)[1]
                default_name = f"{input_name}（脱敏）{input_ext}"
                self.output_file_path = os.path.join(input_dir, default_name)
            
            if self.input_file_path.endswith(('.xlsx', '.xls')):
                self.save_excel_changes()
            elif self.input_file_path.endswith('.txt'):
                self.save_text_changes()
            elif self.input_file_path.endswith('.doc'):
                QMessageBox.information(self, "DOC文件格式提示", "由于兼容性原因，暂不支持直接处理DOC格式文件。\n请先将DOC文件转换为DOCX格式后再进行脱敏处理。")
                return
            elif self.input_file_path.endswith('.docx'):
                self.save_word_changes()
            else:
                QMessageBox.warning(self, "警告", "当前文件类型不支持交互式脱敏")
                return
            
            QMessageBox.information(self, "完成", f"文件已保存到: {self.output_file_path}")

            # 新增：询问是否打开文件
            reply = QMessageBox.question(self, "打开文件", "是否立即打开刚保存的文件？", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                try:
                    import subprocess
                    import platform
                    file_path = self.output_file_path
                    if platform.system() == "Windows":
                        os.startfile(file_path)
                    elif platform.system() == "Darwin":
                        subprocess.Popen(["open", file_path])
                    else:
                        subprocess.Popen(["xdg-open", file_path])
                except Exception as e_open:
                    QMessageBox.critical(self, "打开失败", f"文件已保存，但打开失败：{str(e_open)}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存文件失败: {str(e)}")
    
    def save_excel_changes(self):
        """保存Excel文件的交互式修改（保持格式）"""
        import openpyxl
        from openpyxl import load_workbook
        
        try:
            # 如果有原始文件路径，直接加载原始工作簿以保持格式
            if hasattr(self, 'original_excel_path') and self.original_excel_path:
                wb_new = load_workbook(self.original_excel_path)
            else:
                # fallback：创建新的工作簿
                wb_new = openpyxl.Workbook()
                if wb_new.active:
                    wb_new.remove(wb_new.active)  # 删除默认工作表
            
            # 从界面上的表格获取数据并更新到工作簿
            if hasattr(self, 'current_workbook') and self.current_workbook:
                # 处理当前显示的工作表
                if hasattr(self, 'current_sheet_name') and self.current_sheet_name:
                    ws_new = wb_new[self.current_sheet_name] if self.current_sheet_name in wb_new.sheetnames else wb_new.active
                    
                    # 用界面上修改后的数据更新工作表，但保持原有格式
                    for row in range(self.table_widget.rowCount()):
                        for col in range(self.table_widget.columnCount()):
                            item = self.table_widget.item(row, col)
                            if item:
                                cell = ws_new.cell(row + 1, col + 1)
                                cell.value = item.text()
                                # 如果有保存的格式信息，应用格式
                                self.apply_cell_format(cell, row, col)
            
            wb_new.save(self.output_file_path)
            
        except Exception as e:
            # 如果格式保持失败，使用简化方法
            print(f"格式保持失败，使用简化保存方法: {str(e)}")
            self.save_excel_changes_simple()
    
    def save_excel_changes_simple(self):
        """简化的Excel保存方法（不保持格式）"""
        import openpyxl
        
        # 创建新的工作簿
        wb_new = openpyxl.Workbook()
        ws_new = wb_new.active
        ws_new.title = getattr(self, 'current_sheet_name', 'Sheet1')
        
        # 从界面表格获取数据
        for row in range(self.table_widget.rowCount()):
            for col in range(self.table_widget.columnCount()):
                item = self.table_widget.item(row, col)
                if item and item.text():
                    ws_new.cell(row + 1, col + 1).value = item.text()
        
        wb_new.save(self.output_file_path)
    
    def save_text_changes(self):
        """保存文本文件的交互式修改"""
        # 获取文本编辑器中的内容
        content = self.text_edit.toPlainText()
        
        # 使用原始编码保存
        encoding = getattr(self, 'original_encoding', 'utf-8')
        with open(self.output_file_path, 'w', encoding=encoding) as f:
            f.write(content)
    
    def save_word_changes(self):
        """保存Word文档的交互式修改（保持原格式）"""
        try:
            if not hasattr(self, 'current_word_doc') or not self.current_word_doc:
                QMessageBox.warning(self, "警告", "没有加载的Word文档")
                return
                
            # 获取编辑器中的内容
            new_content = self.word_edit.toPlainText()
            
            # 重新加载原始文档以进行替换操作
            from docx import Document
            
            # 只支持DOCX文件的保存
            if self.input_file_path.lower().endswith('.doc'):
                QMessageBox.warning(self, "不支持的格式", 
                    "DOC格式不支持交互式脱敏保存\n\n请先将DOC文件转换为DOCX格式：\n"
                    "1. 用Word打开DOC文件\n"
                    "2. 另存为DOCX格式\n"
                    "3. 重新选择DOCX文件进行脱敏")
                return
            else:
                # 直接打开DOCX文件
                doc = Document(self.input_file_path)
            
            # 获取原始文本内容（用于比对）
            original_text = self.get_word_text_content(doc)
            
            # 计算需要替换的内容
            replacements = self.calculate_text_replacements(original_text, new_content)
            
            # 应用替换到文档（保持格式）
            self.apply_word_replacements(doc, replacements)
            
            # 保存文档
            if self.output_file_path.lower().endswith(('.doc', '.docx')):
                output_path = self.output_file_path
                if not output_path.lower().endswith('.docx'):
                    output_path = os.path.splitext(output_path)[0] + '.docx'
            else:
                output_path = os.path.splitext(self.output_file_path)[0] + '.docx'
            
            doc.save(output_path)
            self.output_file_path = output_path
            
            # 已保存，主流程统一弹窗，无需此处弹窗
            
        except ImportError:
            QMessageBox.warning(self, "警告", "未安装python-docx库，无法保存DOCX文件")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存Word文档时出错: {str(e)}")

    def get_word_text_content(self, doc):
        """提取Word文档的纯文本内容"""
        text_parts = []
        for para in doc.paragraphs:
            text_parts.append(para.text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text_parts.append(cell.text)
        return '\n'.join(text_parts)
    
    def calculate_text_replacements(self, original_text, new_text):
        """计算文本差异，返回需要替换的内容"""
        # 简单的逐行比较替换策略
        original_lines = original_text.split('\n')
        new_lines = new_text.split('\n')
        
        replacements = []
        min_len = min(len(original_lines), len(new_lines))
        
        for i in range(min_len):
            if original_lines[i] != new_lines[i]:
                replacements.append((original_lines[i], new_lines[i]))
        
        return replacements
    
    def apply_word_replacements(self, doc, replacements):
        """将替换应用到Word文档，保持格式"""
        for old_text, new_text in replacements:
            if old_text.strip() and old_text != new_text:
                # 替换段落中的文本
                for para in doc.paragraphs:
                    if old_text in para.text:
                        self.replace_text_in_paragraph(para, old_text, new_text)
                
                # 替换表格中的文本
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            if old_text in cell.text:
                                for para in cell.paragraphs:
                                    if old_text in para.text:
                                        self.replace_text_in_paragraph(para, old_text, new_text)
    
    def replace_text_in_paragraph(self, paragraph, old_text, new_text):
        """在段落中替换文本，保持格式"""
        if old_text in paragraph.text:
            # 获取段落的所有runs
            runs = paragraph.runs
            paragraph_text = paragraph.text
            
            if old_text in paragraph_text:
                # 找到替换位置
                start_pos = paragraph_text.find(old_text)
                end_pos = start_pos + len(old_text)
                
                # 记录格式信息
                current_pos = 0
                start_run_idx = -1
                end_run_idx = -1
                start_run_pos = 0
                end_run_pos = 0
                
                # 找到起始和结束run的位置
                for i, run in enumerate(runs):
                    run_len = len(run.text)
                    if start_run_idx == -1 and current_pos + run_len > start_pos:
                        start_run_idx = i
                        start_run_pos = start_pos - current_pos
                    if current_pos + run_len >= end_pos:
                        end_run_idx = i
                        end_run_pos = end_pos - current_pos
                        break
                    current_pos += run_len
                
                # 执行替换
                if start_run_idx != -1 and end_run_idx != -1:
                    if start_run_idx == end_run_idx:
                        # 在同一个run内替换
                        run = runs[start_run_idx]
                        run.text = run.text[:start_run_pos] + new_text + run.text[end_run_pos:]
                    else:
                        # 跨多个run替换
                        runs[start_run_idx].text = runs[start_run_idx].text[:start_run_pos] + new_text
                        for i in range(start_run_idx + 1, end_run_idx + 1):
                            if i < len(runs):
                                if i == end_run_idx:
                                    runs[i].text = runs[i].text[end_run_pos:]
                                else:
                                    runs[i].text = ""

    def auto_process_file(self):
        """自动脱敏模式 - 应用规则到整个文件"""
            
        try:
            # 确保输出目录存在
            os.makedirs(os.path.dirname(self.output_file_path), exist_ok=True)

            # Excel
            if self.input_file_path.lower().endswith('.xlsx'):
                try:
                    from openpyxl import load_workbook
                    from copy import copy
                    import re

                    wb_original = load_workbook(self.input_file_path)
                    wb_new = load_workbook(self.input_file_path)  # 创建副本保留格式

                    # 记录自动脱敏的历史
                    auto_redaction_history = []

                    for ws_name in wb_original.sheetnames:
                        ws_original = wb_original[ws_name]
                        ws_new = wb_new[ws_name]

                        # 复制合并单元格信息
                        try:
                            for merged_range in list(ws_original.merged_cells.ranges):
                                ws_new.merge_cells(str(merged_range))
                        except:
                            pass  # 如果合并单元格操作失败，继续处理

                        for row in ws_original.iter_rows():
                            for cell in row:
                                if cell.value is not None:
                                    cell_value = str(cell.value)
                                    original_value = cell_value
                                    for rule in self.rule_engine.get_active_rules():
                                        if rule.name == "自定义字段":
                                            custom_fields = getattr(self, 'custom_fields', None)
                                            cell_value = self.rule_engine.apply_redaction_rule(rule, cell_value, custom_fields)
                                        else:
                                            custom_names = getattr(self, 'custom_names', None)
                                            cell_value = self.rule_engine.apply_redaction_rule(rule, cell_value, custom_names)
                                    
                                    # 更新单元格值并保持格式
                                    target_cell = ws_new[cell.coordinate]
                                    if cell_value != original_value:
                                        target_cell.value = cell_value
                                        
                                        # 记录自动脱敏历史（仅限活动工作表用于界面显示）
                                        if ws_name == wb_original.active.title:
                                            auto_redaction_history.append({
                                                'row': cell.row - 1,  # 转换为0索引
                                                'col': cell.column - 1,
                                                'original_text': original_value,
                                                'redacted_text': cell_value,
                                                'original_background': QColor(),
                                                'original_tooltip': ''
                                            })
                                    
                                    # 复制所有格式属性
                                    if cell.has_style:
                                        target_cell.font = copy(cell.font)
                                        target_cell.border = copy(cell.border)
                                        target_cell.fill = copy(cell.fill)
                                        target_cell.number_format = cell.number_format
                                        target_cell.protection = copy(cell.protection)
                                        target_cell.alignment = copy(cell.alignment)

                    wb_new.save(self.output_file_path)
                    
                    # 保存历史记录
                    if auto_redaction_history:
                        self.excel_redaction_history.append({
                            'type': 'auto_rule_redaction',
                            'operations': auto_redaction_history
                        })
                    
                    # 在界面中显示脱敏后的结果（优化输出格式与交互模式一致）
                    try:
                        # 保存工作簿信息以便后续操作
                        self.current_workbook = wb_new
                        self.current_sheet_name = wb_new.active.title
                        
                        ws = wb_new.active
                        
                        # 清空表格
                        self.table_widget.clear()
                        
                        # 获取行数和列数
                        max_row = ws.max_row if hasattr(ws, 'max_row') else 0
                        max_col = ws.max_column if hasattr(ws, 'max_column') else 0
                        self.table_widget.setRowCount(max_row)
                        self.table_widget.setColumnCount(max_col)
                        
                        # 加载脱敏后的Excel数据到界面
                        for row_idx, row in enumerate(ws.iter_rows(), 1):
                            for col_idx, cell in enumerate(row, 1):
                                if cell and cell.value is not None:
                                    try:
                                        value = str(cell.value) if cell.value is not None else ""
                                        table_item = QTableWidgetItem(value)
                                        
                                        # 检查是否为脱敏单元格，如果是则标记
                                        for operation in auto_redaction_history:
                                            if (operation['row'] == row_idx - 1 and 
                                                operation['col'] == col_idx - 1):
                                                table_item.setBackground(QColor(255, 235, 235))  # 浅红色背景
                                                table_item.setToolTip(f"已脱敏 - 原文本: {operation['original_text']}")
                                                break
                                        
                                        self.table_widget.setItem(
                                            row_idx - 1,  # QTableWidget从0开始
                                            col_idx - 1,
                                            table_item)
                                    except Exception:
                                        continue
                        
                        # 切换到Excel标签页显示结果
                        self.content_tabs.setCurrentIndex(1)
                        
                    except Exception as e:
                        QMessageBox.warning(self, "警告", f"在界面显示脱敏结果失败: {str(e)}")

                except Exception as e:
                    QMessageBox.warning(self, "警告", f"处理Excel文件失败: {str(e)}")
                    return

            # Word (.docx 或 .doc)
            elif self.input_file_path.lower().endswith(('.docx', '.doc')):
                try:
                    # 仅提示 .doc，让用户先转换
                    if self.input_file_path.lower().endswith('.doc'):
                        QMessageBox.information(
                            self,
                            "格式提示",
                            "很抱歉，由于兼容性原因，暂不支持直接处理DOC格式文件。\n请先使用word将DOC文件转换为DOCX格式后再继续脱敏工作。"

                        )
                        return

                    # 检查文件是否被占用
                    try:
                        with open(self.input_file_path, 'r+b') as test_file:
                            pass  # 如果能打开，说明文件未被占用
                    except PermissionError:
                        QMessageBox.critical(self, "文件被占用", 
                                           f"无法处理Word文档，文件可能正在被使用：\n\n"
                                           f"请关闭可能正在使用该文件的程序（如Microsoft Word）后重试。\n\n"
                                           f"文件路径：{self.input_file_path}")
                        return

                    from docx import Document
                    doc = Document(self.input_file_path)

                    # 段落
                    for para in doc.paragraphs:
                        original_text = para.text
                        processed_text = original_text
                        for rule in self.rule_engine.get_active_rules():
                            if rule.name == "自定义字段":
                                custom_fields = getattr(self, 'custom_fields', None)
                                processed_text = self.rule_engine.apply_redaction_rule(rule, processed_text, custom_fields)
                            else:
                                custom_names = getattr(self, 'custom_names', None)
                                processed_text = self.rule_engine.apply_redaction_rule(rule, processed_text, custom_names)
                        if original_text != processed_text:
                            para.text = processed_text

                    # 表格
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                original_text = cell.text
                                processed_text = original_text
                                for rule in self.rule_engine.get_active_rules():
                                    if rule.name == "自定义字段":
                                        custom_fields = getattr(self, 'custom_fields', None)
                                        processed_text = self.rule_engine.apply_redaction_rule(rule, processed_text, custom_fields)
                                    else:
                                        custom_names = getattr(self, 'custom_names', None)
                                        processed_text = self.rule_engine.apply_redaction_rule(rule, processed_text, custom_names)
                                if original_text != processed_text:
                                    cell.text = processed_text

                    # 保存为 docx
                    if self.output_file_path.lower().endswith('.docx'):
                        doc.save(self.output_file_path)
                    else:
                        output_path = os.path.splitext(self.output_file_path)[0] + '.docx'
                        doc.save(output_path)
                        self.output_file_path = output_path

                except ImportError:
                    QMessageBox.warning(self, "警告", "处理Word文档需要安装python-docx库\n请运行: pip install python-docx")
                    return
                except PermissionError as e:
                    QMessageBox.critical(self, "权限错误", 
                                       f"无法访问Word文档，请检查：\n\n"
                                       f"1. 是否有其他程序（如Microsoft Word）正在使用该文件\n"
                                       f"2. 文件是否被设置为只读\n"
                                       f"3. 是否有足够的文件访问权限\n\n"
                                       f"错误详情: {str(e)}")
                    return
                except Exception as e:
                    error_msg = str(e)
                    if "Permission denied" in error_msg or "Errno 13" in error_msg:
                        QMessageBox.critical(self, "权限错误", 
                                           f"无法访问Word文档，请检查：\n\n"
                                           f"1. 是否有其他程序（如Microsoft Word）正在使用该文件\n"
                                           f"2. 文件是否被设置为只读\n"
                                           f"3. 是否有足够的文件访问权限\n\n"
                                           f"错误详情: {error_msg}")
                    else:
                        QMessageBox.warning(self, "警告", f"处理Word文档失败: {error_msg}")
                    return

            # 文本
            elif self.input_file_path.lower().endswith('.txt'):
                encoding = None
                encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig', 'latin1']
                import re

                lines = []
                for enc in encodings:
                    try:
                        with open(self.input_file_path, 'r', encoding=enc) as f:
                            lines = f.readlines()
                        encoding = enc
                        break
                    except UnicodeDecodeError:
                        continue

                if encoding is None or not lines:
                    QMessageBox.critical(self, "错误", "无法读取文件内容或确定编码格式")
                    return

                processed_lines = []
                for line in lines:
                    processed_line = line
                    for rule in self.rule_engine.get_active_rules():
                        if rule.name == "自定义字段":
                            custom_fields = getattr(self, 'custom_fields', None)
                            processed_line = self.rule_engine.apply_redaction_rule(rule, processed_line, custom_fields)
                        else:
                            custom_names = getattr(self, 'custom_names', None)
                            processed_line = self.rule_engine.apply_redaction_rule(rule, processed_line, custom_names)
                    processed_lines.append(processed_line)

                with open(self.output_file_path, 'w', encoding=encoding) as f:
                    f.writelines(processed_lines)

            else:
                QMessageBox.warning(self, "警告", "不支持的文件格式")

            QMessageBox.information(self, "成功", "文件脱敏处理完成")
            self.status_label.setText("处理完成")

            # 新增：询问是否打开文件
            reply = QMessageBox.question(self, "打开文件", "是否立即打开刚保存的文件？", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                try:
                    import subprocess
                    import platform
                    file_path = self.output_file_path
                    if platform.system() == "Windows":
                        os.startfile(file_path)
                    elif platform.system() == "Darwin":
                        subprocess.Popen(["open", file_path])
                    else:
                        subprocess.Popen(["xdg-open", file_path])
                except Exception as e_open:
                    QMessageBox.critical(self, "打开失败", f"文件已保存，但打开失败：{str(e_open)}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理文件失败: {str(e)}")
            self.status_label.setText("处理失败")

    def batch_process(self):
        """批量处理文件"""
        # 第一步：检查是否在自动规则模式
        if self.mode_combo.currentIndex() != 1:
            QMessageBox.warning(self, "提示", "批量处理功能仅在【自动脱敏（规则模式）】下可用！\n\n请先切换到自动脱敏模式。")
            return
        
        # 第二步：检查是否有激活的规则
        active_rules = self.rule_engine.get_active_rules()
        if not active_rules:
            QMessageBox.warning(self, "提示", "没有激活的脱敏规则！\n\n请先点击【配置脱敏规则】按钮激活需要的规则。")
            return
        
        # 第三步：显示步骤提示和选择处理方式
        help_msg = """
        <h3>📦 批量处理使用说明</h3>
        
        <p><b>✅ 当前激活规则：</b><br>
        """ + "<br>".join([f"• {rule.name}" for rule in active_rules]) + """</p>
        
        <p><b>📁 支持格式：</b>Excel (.xlsx)、Word (.docx)、文本 (.txt)</p>
        
        <p><b>� 使用步骤：</b>选择文件或文件夹 → 选择输出目录 → 自动处理完成</p>
        
        <p><b>�💡 提示：</b>输出文件将自动添加"（脱敏）"标识</p>
        """
        
        # 选择批量处理方式
        choice_dialog = QMessageBox(self)
        choice_dialog.setWindowTitle("批量处理方式选择")
        choice_dialog.setText(help_msg + "\n\n请选择批量处理方式：")
        choice_dialog.setIcon(QMessageBox.Icon.Question)
        
        folder_btn = choice_dialog.addButton("📁 选择文件夹", QMessageBox.ButtonRole.AcceptRole)
        multi_files_btn = choice_dialog.addButton("📄 多选文件", QMessageBox.ButtonRole.AcceptRole)
        cancel_btn = choice_dialog.addButton("❌ 取消", QMessageBox.ButtonRole.RejectRole)
        
        choice_dialog.exec()
        clicked_btn = choice_dialog.clickedButton()
        
        if clicked_btn == cancel_btn:
            return
            
        # 根据选择方式获取文件列表
        input_files = []
        output_dir = None
        
        if clicked_btn == folder_btn:
            # 选择文件夹方式
            default_dir = os.path.dirname(os.path.abspath(__file__))
            input_dir = QFileDialog.getExistingDirectory(
                self,
                "第1步：选择待处理文件夹",
                default_dir
            )
            if not input_dir:
                return
                
            # 选择输出文件夹
            output_dir = QFileDialog.getExistingDirectory(
                self,
                "第2步：选择输出目录",
                default_dir
            )
            if not output_dir:
                return
            
            # 获取文件夹中的所有支持文件
            for filename in os.listdir(input_dir):
                if filename.lower().endswith(('.xlsx', '.docx', '.txt')):
                    input_files.append(os.path.join(input_dir, filename))
                    
        elif clicked_btn == multi_files_btn:
            # 多选文件方式
            script_dir = os.path.dirname(os.path.abspath(__file__))
            file_paths, _ = QFileDialog.getOpenFileNames(
                self,
                "选择需要批量处理的文件（可多选）",
                script_dir,
                "支持的文件 (*.xlsx *.docx *.txt);;Excel文件 (*.xlsx);;Word文档 (*.docx);;文本文件 (*.txt);;所有文件 (*)"
            )
            if not file_paths:
                return
                
            input_files = file_paths
            
            # 选择输出文件夹
            default_dir = os.path.dirname(os.path.abspath(__file__))
            output_dir = QFileDialog.getExistingDirectory(
                self,
                "选择输出目录",
                default_dir
            )
            if not output_dir:
                return
        
        if not input_files:
            QMessageBox.information(self, "提示", "没有找到支持的文件格式\n\n支持格式：.xlsx、.docx、.txt")
            return
        
        # 确保 output_dir 不为 None
        if output_dir is None:
            QMessageBox.warning(self, "错误", "输出目录未选择")
            return
        
        try:
            processed_count = 0
            failed_files = []
            
            self.progress_bar.setRange(0, len(input_files))
            self.progress_bar.setValue(0)
            
            for i, input_path in enumerate(input_files):
                filename = os.path.basename(input_path)
                
                # 生成带（脱敏）标识的输出文件名
                name_without_ext = os.path.splitext(filename)[0]
                file_ext = os.path.splitext(filename)[1]
                output_filename = f"{name_without_ext}（脱敏）{file_ext}"
                output_path = os.path.join(output_dir, output_filename)
                
                if not os.path.isfile(input_path):
                    continue
                    
                try:
                    # 处理Excel文件
                    if filename.lower().endswith('.xlsx'):
                        self.status_label.setText(f"正在处理Excel: {filename}")
                        QApplication.processEvents()
                        
                        try:
                            from openpyxl import load_workbook
                            from copy import copy
                            
                            wb_original = load_workbook(input_path)
                            wb_new = load_workbook(input_path)  # 创建副本保留格式
                            
                            for ws_name in wb_original.sheetnames:
                                ws_original = wb_original[ws_name]
                                ws_new = wb_new[ws_name]
                                
                                for row in ws_original.iter_rows():
                                    for cell in row:
                                        if cell.value is not None:
                                            cell_value = str(cell.value)
                                            original_value = cell_value  # 保存原始值
                                            
                                            # 使用规则引擎进行脱敏处理
                                            for rule in active_rules:
                                                if rule.name == "姓名":
                                                    custom_names = getattr(self, 'custom_names', None)
                                                    cell_value = self.rule_engine.apply_redaction_rule(rule, cell_value, custom_names)
                                                elif rule.name == "自定义字段":
                                                    custom_fields = getattr(self, 'custom_fields', None)
                                                    cell_value = self.rule_engine.apply_redaction_rule(rule, cell_value, custom_fields)
                                                else:
                                                    cell_value = self.rule_engine.apply_redaction_rule(rule, cell_value)
                                            
                                            if cell_value != original_value:  # 仅在值发生变化时更新
                                                ws_new[cell.coordinate].value = cell_value
                                            if hasattr(cell, 'style'):
                                                ws_new[cell.coordinate].style = copy(cell.style)
                            
                            wb_new.save(output_path)
                            
                        except Exception as e:
                            failed_files.append(f"{filename} (Excel处理失败: {str(e)})")
                            continue
                            
                    elif filename.lower().endswith('.docx'):
                        # 处理Word文档
                        self.status_label.setText(f"正在处理Word: {filename}")
                        QApplication.processEvents()
                        
                        try:
                            from docx import Document
                            
                            doc = Document(input_path)
                            
                            # 处理段落文本
                            for para in doc.paragraphs:
                                original_text = para.text
                                processed_text = original_text
                                
                                # 应用所有激活的脱敏规则
                                for rule in active_rules:
                                    if rule.name == "姓名":
                                        custom_names = getattr(self, 'custom_names', None)
                                        processed_text = self.rule_engine.apply_redaction_rule(rule, processed_text, custom_names)
                                    elif rule.name == "自定义字段":
                                        custom_fields = getattr(self, 'custom_fields', None)
                                        processed_text = self.rule_engine.apply_redaction_rule(rule, processed_text, custom_fields)
                                    else:
                                        processed_text = self.rule_engine.apply_redaction_rule(rule, processed_text)
                                
                                # 更新段落文本
                                if original_text != processed_text:
                                    para.text = processed_text
                            
                            # 处理表格中的文本
                            for table in doc.tables:
                                for row in table.rows:
                                    for cell in row.cells:
                                        original_text = cell.text
                                        processed_text = original_text
                                        
                                        # 应用所有激活的脱敏规则
                                        for rule in active_rules:
                                            if rule.name == "姓名":
                                                custom_names = getattr(self, 'custom_names', None)
                                                processed_text = self.rule_engine.apply_redaction_rule(rule, processed_text, custom_names)
                                            elif rule.name == "自定义字段":
                                                custom_fields = getattr(self, 'custom_fields', None)
                                                processed_text = self.rule_engine.apply_redaction_rule(rule, processed_text, custom_fields)
                                            else:
                                                processed_text = self.rule_engine.apply_redaction_rule(rule, processed_text)
                                        
                                        # 更新单元格文本
                                        if original_text != processed_text:
                                            cell.text = processed_text
                            
                            # 保存DOCX文档
                            doc.save(output_path)
                            
                        except ImportError:
                            failed_files.append(f"{filename} (需要python-docx库处理Word文档)")
                            continue
                        except Exception as e:
                            failed_files.append(f"{filename} (Word处理失败: {str(e)})")
                            continue
                            
                    elif filename.lower().endswith('.txt'):
                        # 处理文本文件
                        self.status_label.setText(f"正在处理文本: {filename}")
                        QApplication.processEvents()
                        
                        encoding = None
                        encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig', 'latin1']
                        content = ""
                        
                        # 读取文件内容
                        for enc in encodings:
                            try:
                                with open(input_path, 'r', encoding=enc) as f:
                                    content = f.read()
                                encoding = enc
                                break
                            except UnicodeDecodeError:
                                continue
                        
                        if encoding is None:
                            failed_files.append(f"{filename} (编码问题)")
                            continue
                        
                        # 使用规则引擎处理内容
                        processed_content = content
                        for rule in active_rules:
                            if rule.name == "姓名":
                                custom_names = getattr(self, 'custom_names', None)
                                processed_content = self.rule_engine.apply_redaction_rule(rule, processed_content, custom_names)
                            elif rule.name == "自定义字段":
                                custom_fields = getattr(self, 'custom_fields', None)
                                processed_content = self.rule_engine.apply_redaction_rule(rule, processed_content, custom_fields)
                            else:
                                processed_content = self.rule_engine.apply_redaction_rule(rule, processed_content)
                        
                        # 写入文件
                        with open(output_path, 'w', encoding=encoding) as f:
                            f.write(processed_content)
                    
                    processed_count += 1
                    self.progress_bar.setValue(i + 1)
                    QApplication.processEvents()
                    
                except Exception as e:
                    failed_files.append(f"{filename} ({str(e)})")
            
            # 显示处理结果
            self.progress_bar.setValue(len(input_files))
            result_msg = f"✅ 批量处理完成！\n\n📊 处理统计：\n• 成功处理：{processed_count} 个文件\n• 输出位置：{output_dir}"
            
            if failed_files:
                result_msg += f"\n• 失败文件：{len(failed_files)} 个"
                if len(failed_files) <= 3:
                    result_msg += f"\n\n❌ 失败详情：\n" + "\n".join([f"• {f}" for f in failed_files])
                else:
                    result_msg += f"\n\n❌ 失败详情（前3个）：\n" + "\n".join([f"• {f}" for f in failed_files[:3]])
                    result_msg += f"\n... 还有 {len(failed_files)-3} 个文件失败"
            
            QMessageBox.information(self, "批量处理完成", result_msg)
            self.status_label.setText("✅ 批量处理完成")
            self.progress_bar.setValue(0)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"批量处理失败: {str(e)}")
            self.status_label.setText("❌ 批量处理失败")
            self.progress_bar.setValue(0)

    def show_text_context_menu(self, position):
        """显示文本编辑器的右键菜单"""
        # 在自动脱敏模式下不显示右键菜单
        if self.mode_combo.currentIndex() == 1:  # 自动脱敏模式
            return
        self.text_menu.exec(self.text_edit.mapToGlobal(position))
        
    def mark_text_redaction(self):
        """标记选中的文本为脱敏内容"""
        cursor = self.text_edit.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            if not selected_text.strip():
                QMessageBox.warning(self, "警告", "选中的文本为空")
                return
                
            # 智能脱敏选中文本
            redacted_text = self.smart_redact_text(selected_text)
            
            # 记录撤销历史
            start_pos = cursor.selectionStart()
            end_pos = cursor.selectionEnd()
            self.text_redaction_history.append({
                'start': start_pos,
                'end': start_pos + len(redacted_text),
                'original': selected_text,
                'redacted': redacted_text
            })
            
            cursor.insertText(redacted_text)
            self.text_edit.setTextCursor(cursor)
        else:
            QMessageBox.warning(self, "警告", "请先选择要脱敏的文本")

    def mark_text_redaction_all(self):
        """标记选中文本在文本编辑器中的所有相同内容为脱敏"""
        cursor = self.text_edit.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            if not selected_text.strip():
                QMessageBox.warning(self, "警告", "选中的文本为空")
                return
                
            # 获取全文内容
            full_text = self.text_edit.toPlainText()
            
            # 智能脱敏选中文本
            redacted_text = self.smart_redact_text(selected_text)
            
            # 记录批量撤销历史
            count = full_text.count(selected_text)
            self.text_redaction_history.append({
                'type': 'replace_all',
                'original': selected_text,
                'redacted': redacted_text,
                'count': count,
                'full_original': full_text
            })
            
            # 在全文中替换所有相同的内容
            new_text = full_text.replace(selected_text, redacted_text)
            
            # 更新文本内容
            self.text_edit.setPlainText(new_text)
            
            # 显示替换结果
            QMessageBox.information(self, "脱敏完成", f"已替换 {count} 处相同内容：\n原文：{selected_text}\n脱敏后：{redacted_text}")
        else:
            QMessageBox.warning(self, "警告", "请先选择要脱敏的文本")
    
    def smart_redact_text(self, text):
        """智能脱敏文本内容"""
        import re
        
        # 检测文本类型并应用相应脱敏规则
        text = text.strip()
        
        # 身份证号（18位）
        if re.match(r'^\d{18}$', text):
            return text[:3] + "*" * 11 + text[-4:]
        
        # 手机号（11位数字）
        elif re.match(r'^1[3-9]\d{9}$', text):
            return text[:3] + "****" + text[-4:]
        
        # 邮箱地址
        elif '@' in text and re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', text):
            at_index = text.find('@')
            return text[0] + "*" * (at_index - 1) + text[at_index:]
        
        # 中文姓名（2-4个汉字）
        elif re.match(r'^[\u4e00-\u9fa5]{2,4}$', text):
            return text[0] + "*" * (len(text) - 1)
        
        # 银行卡号（16-19位数字）
        elif re.match(r'^\d{16,19}$', text):
            return text[:4] + "*" * (len(text) - 8) + text[-4:]
        
        # 默认脱敏：保留首尾，中间用*替换
        else:
            if len(text) <= 2:
                return "*" * len(text)
        
        # 对于其他任意文本，简单保留首尾字符
        if len(text) > 2:
            return text[0] + ("*" * (len(text) - 2)) + text[-1]
        return text

    def show_word_context_menu(self, position):
        """显示Word编辑器的右键菜单"""
        # 在自动脱敏模式下不显示右键菜单
        if self.mode_combo.currentIndex() == 1:  # 自动脱敏模式
            return
        self.word_menu.exec(self.word_edit.mapToGlobal(position))
        
    def mark_word_redaction(self):
        """标记选中的Word文档文本为脱敏内容"""
        cursor = self.word_edit.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            if not selected_text.strip():
                QMessageBox.warning(self, "警告", "选中的文本为空")
                return
                
            # 智能脱敏选中文本
            redacted_text = self.smart_redact_text(selected_text)
            
            # 记录撤销历史
            start_pos = cursor.selectionStart()
            end_pos = cursor.selectionEnd()
            self.word_redaction_history.append({
                'start': start_pos,
                'end': start_pos + len(redacted_text),
                'original': selected_text,
                'redacted': redacted_text
            })
            
            cursor.insertText(redacted_text)
            self.word_edit.setTextCursor(cursor)
        else:
            QMessageBox.warning(self, "警告", "请先选择要脱敏的文本")

    def mark_word_redaction_all(self):
        """标记选中文本在Word文档中的所有相同内容为脱敏"""
        cursor = self.word_edit.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            if not selected_text.strip():
                QMessageBox.warning(self, "警告", "选中的文本为空")
                return
                
            # 获取全文内容
            full_text = self.word_edit.toPlainText()
            
            # 智能脱敏选中文本
            redacted_text = self.smart_redact_text(selected_text)
            
            # 记录批量撤销历史
            count = full_text.count(selected_text)
            self.word_redaction_history.append({
                'type': 'replace_all',
                'original': selected_text,
                'redacted': redacted_text,
                'count': count,
                'full_original': full_text
            })
            
            # 在全文中替换所有相同的内容
            new_text = full_text.replace(selected_text, redacted_text)
            
            # 更新文本内容
            self.word_edit.setPlainText(new_text)
            
            # 显示替换结果
            QMessageBox.information(self, "脱敏完成", f"已替换 {count} 处相同内容：\n原文：{selected_text}\n脱敏后：{redacted_text}")
        else:
            QMessageBox.warning(self, "警告", "请先选择要脱敏的文本")

    def setup_table_context_menu(self):
        """设置表格右键菜单"""
        self.table_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table_widget.customContextMenuRequested.connect(self.show_table_context_menu)
        
        self.table_menu = QMenu(self)
        # 添加菜单项
        self.redact_cell_action = QAction("脱敏选中单元格", self)
        self.redact_cell_action.triggered.connect(self.mark_cell_redaction)
        self.table_menu.addAction(self.redact_cell_action)
        
        self.redact_row_action = QAction("脱敏整行", self)
        self.redact_row_action.triggered.connect(self.mark_row_redaction)
        self.table_menu.addAction(self.redact_row_action)
        
        self.redact_col_action = QAction("脱敏整列", self)
        self.redact_col_action.triggered.connect(self.mark_column_redaction)
        self.table_menu.addAction(self.redact_col_action)
        
        # 添加全局查找替换脱敏功能
        self.table_redact_all_action = QAction("🔄 全表相同内容脱敏", self)
        self.table_redact_all_action.triggered.connect(self.mark_table_redaction_all)
        self.table_menu.addAction(self.table_redact_all_action)
        
        # 添加分隔符
        self.table_menu.addSeparator()
        
        # 添加撤销脱敏功能
        self.undo_redaction_action = QAction("撤销脱敏", self)
        self.undo_redaction_action.triggered.connect(self.undo_redaction)
        self.table_menu.addAction(self.undo_redaction_action)
        
        # 添加撤销当前区域脱敏功能
        self.excel_undo_current_action = QAction("撤销当前区域脱敏", self)
        self.excel_undo_current_action.triggered.connect(self.undo_current_excel_redaction)
        self.table_menu.addAction(self.excel_undo_current_action)

    def show_table_context_menu(self, position):
        """显示表格的右键菜单"""
        # 在自动脱敏模式下不显示右键菜单
        if self.mode_combo.currentIndex() == 1:  # 自动脱敏模式
            return
        
        # 保存右键点击的位置对应的单元格索引，用于整行/整列脱敏
        item = self.table_widget.itemAt(position)
        if item:
            self.current_right_click_row = item.row()
            self.current_right_click_col = item.column()
        else:
            self.current_right_click_row = -1
            self.current_right_click_col = -1
        self.table_menu.exec(self.table_widget.mapToGlobal(position))

    def mark_cell_redaction(self):
        """标记选中的单元格为脱敏内容"""
        selected_items = self.table_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先选择要脱敏的单元格")
            return
            
        # 获取用户确认
        reply = QMessageBox.question(self, "确认脱敏", 
                                   f"确定要对选中的 {len(selected_items)} 个单元格进行脱敏吗？\n请慎重操作！",
                                   QMessageBox.StandardButton.Yes, QMessageBox.StandardButton.No)
        
        if reply != QMessageBox.StandardButton.Yes:
            return
            
        # 记录脱敏的单元格数量和历史记录
        redacted_count = 0
        operation_history = []
        
        for item in selected_items:
            if item and item.text().strip():
                original_text = item.text().strip()
                redacted_text = self.smart_redact_text(original_text)
                if redacted_text != original_text:
                    # 记录历史
                    operation_history.append({
                        'row': item.row(),
                        'col': item.column(),
                        'original_text': original_text,
                        'redacted_text': redacted_text,
                        'original_background': item.background(),
                        'original_tooltip': item.toolTip()
                    })
                    
                    item.setText(redacted_text)
                    # 标记脱敏的单元格
                    item.setBackground(QColor(255, 235, 235))  # 浅红色背景
                    item.setToolTip(f"已脱敏 - 原文本: {original_text}")
                    redacted_count += 1
        
        # 记录撤销历史
        if operation_history:
            self.excel_redaction_history.append({
                'type': 'cell_redaction',
                'operations': operation_history
            })
        
        if redacted_count > 0:
            QMessageBox.information(self, "脱敏完成", f"已成功脱敏 {redacted_count} 个单元格")
        else:
            QMessageBox.information(self, "提示", "没有找到需要脱敏的内容")

    def mark_row_redaction(self):
        """标记整行为脱敏内容（基于右键点击的单元格位置）"""
        # 如果有右键点击的单元格，则脱敏该行
        if hasattr(self, 'current_right_click_row') and self.current_right_click_row >= 0:
            target_row = self.current_right_click_row
            total_cells = self.table_widget.columnCount()
            
            # 获取用户确认
            effective_cells = (total_cells - 1) if total_cells > 1 else total_cells
            reply = QMessageBox.question(self, "确认脱敏", 
                                       f"确定要对第 {target_row + 1} 行（{effective_cells} 个单元格，跳过第1列）进行脱敏吗？\n请慎重操作！",
                                       QMessageBox.StandardButton.Yes, QMessageBox.StandardButton.No)
            
            if reply != QMessageBox.StandardButton.Yes:
                return
                
            # 记录脱敏的单元格数量和历史记录
            redacted_count = 0
            operation_history = []
            
            # 遍历该行的所有列（跳过第一列，通常是序号列）
            start_col = 1 if self.table_widget.columnCount() > 1 else 0
            for col in range(start_col, self.table_widget.columnCount()):
                item = self.table_widget.item(target_row, col)
                if item and item.text().strip():
                    original_text = item.text().strip()
                    redacted_text = self.smart_redact_text(original_text)
                    if redacted_text != original_text:
                        # 记录历史
                        operation_history.append({
                            'row': target_row,
                            'col': col,
                            'original_text': original_text,
                            'redacted_text': redacted_text,
                            'original_background': item.background(),
                            'original_tooltip': item.toolTip()
                        })
                        
                        item.setText(redacted_text)
                        # 标记脱敏的单元格
                        item.setBackground(QColor(255, 235, 235))  # 浅红色背景
                        item.setToolTip(f"已脱敏 - 原文本: {original_text}")
                        redacted_count += 1
            
            # 记录撤销历史
            if operation_history:
                self.excel_redaction_history.append({
                    'type': 'row_redaction',
                    'operations': operation_history
                })
            
            if redacted_count > 0:
                QMessageBox.information(self, "脱敏完成", f"已成功脱敏第 {target_row + 1} 行的 {redacted_count} 个单元格")
            else:
                QMessageBox.information(self, "提示", f"第 {target_row + 1} 行没有找到需要脱敏的内容")
        else:
            # 回退到原来的逻辑：处理用户选中的范围
            selected_ranges = self.table_widget.selectedRanges()
            if not selected_ranges:
                QMessageBox.warning(self, "警告", "请先点击要脱敏的单元格")
                return
                
            # 计算总单元格数量
            total_cells = 0
            for r in selected_ranges:
                total_cells += (r.bottomRow() - r.topRow() + 1) * self.table_widget.columnCount()
                
            # 获取用户确认
            reply = QMessageBox.question(self, "确认脱敏", 
                                       f"确定要对选中的行（约 {total_cells} 个单元格）进行脱敏吗？\n请慎重操作！",
                                       QMessageBox.StandardButton.Yes, QMessageBox.StandardButton.No)
            
            if reply != QMessageBox.StandardButton.Yes:
                return
                
            # 记录脱敏的单元格数量和历史记录
            redacted_count = 0
            operation_history = []
            
            for r in selected_ranges:
                for row in range(r.topRow(), r.bottomRow() + 1):
                    for col in range(self.table_widget.columnCount()):
                        item = self.table_widget.item(row, col)
                        if item and item.text().strip():
                            original_text = item.text().strip()
                            redacted_text = self.smart_redact_text(original_text)
                            if redacted_text != original_text:
                                # 记录历史
                                operation_history.append({
                                    'row': row,
                                    'col': col,
                                    'original_text': original_text,
                                    'redacted_text': redacted_text,
                                    'original_background': item.background(),
                                    'original_tooltip': item.toolTip()
                                })
                                
                                item.setText(redacted_text)
                                # 标记脱敏的单元格
                                item.setBackground(QColor(255, 235, 235))  # 浅红色背景
                                item.setToolTip(f"已脱敏 - 原文本: {original_text}")
                                redacted_count += 1
            
            # 记录撤销历史
            if operation_history:
                self.excel_redaction_history.append({
                    'type': 'row_range_redaction',
                    'operations': operation_history
                })
            
            if redacted_count > 0:
                QMessageBox.information(self, "脱敏完成", f"已成功脱敏 {redacted_count} 个单元格")
            else:
                QMessageBox.information(self, "提示", "没有找到需要脱敏的内容")

    def mark_column_redaction(self):
        """标记整列为脱敏内容（基于右键点击的单元格位置）"""
        # 如果有右键点击的单元格，则脱敏该列
        if hasattr(self, 'current_right_click_col') and self.current_right_click_col >= 0:
            target_col = self.current_right_click_col
            total_cells = self.table_widget.rowCount()
            
            # 获取用户确认
            effective_cells = (total_cells - 1) if total_cells > 1 else total_cells
            reply = QMessageBox.question(self, "确认脱敏", 
                                       f"确定要对第 {target_col + 1} 列（{effective_cells} 个单元格，跳过第1行）进行脱敏吗？\n请慎重操作！",
                                       QMessageBox.StandardButton.Yes, QMessageBox.StandardButton.No)
            
            if reply != QMessageBox.StandardButton.Yes:
                return
                
            # 记录脱敏的单元格数量和历史记录
            redacted_count = 0
            operation_history = []
            
            # 遍历该列的所有行（跳过第一行，通常是表头）
            start_row = 1 if self.table_widget.rowCount() > 1 else 0
            for row in range(start_row, self.table_widget.rowCount()):
                item = self.table_widget.item(row, target_col)
                if item and item.text().strip():
                    original_text = item.text().strip()
                    redacted_text = self.smart_redact_text(original_text)
                    if redacted_text != original_text:
                        # 记录历史
                        operation_history.append({
                            'row': row,
                            'col': target_col,
                            'original_text': original_text,
                            'redacted_text': redacted_text,
                            'original_background': item.background(),
                            'original_tooltip': item.toolTip()
                        })
                        
                        item.setText(redacted_text)
                        # 标记脱敏的单元格
                        item.setBackground(QColor(255, 235, 235))  # 浅红色背景
                        item.setToolTip(f"已脱敏 - 原文本: {original_text}")
                        redacted_count += 1
            
            # 记录撤销历史
            if operation_history:
                self.excel_redaction_history.append({
                    'type': 'column_redaction',
                    'operations': operation_history
                })
            
            if redacted_count > 0:
                QMessageBox.information(self, "脱敏完成", f"已成功脱敏第 {target_col + 1} 列的 {redacted_count} 个单元格")
            else:
                QMessageBox.information(self, "提示", f"第 {target_col + 1} 列没有找到需要脱敏的内容")
        else:
            # 回退到原来的逻辑：处理用户选中的范围
            selected_ranges = self.table_widget.selectedRanges()
            if not selected_ranges:
                QMessageBox.warning(self, "警告", "请先点击要脱敏的单元格")
                return
                
            # 获取用户确认
            reply = QMessageBox.question(self, "确认脱敏", 
                                       "确定要对选中的单元格进行脱敏吗？\n请慎重操作！",
                                       QMessageBox.StandardButton.Yes, QMessageBox.StandardButton.No)
            
            if reply != QMessageBox.StandardButton.Yes:
                return
                
            # 记录脱敏的单元格数量和历史记录
            redacted_count = 0
            operation_history = []
            
            # 只处理选中的单元格范围
            for r in selected_ranges:
                for row in range(r.topRow(), r.bottomRow() + 1):
                    for col in range(r.leftColumn(), r.rightColumn() + 1):
                        item = self.table_widget.item(row, col)
                        if item and item.text().strip():
                            original_text = item.text().strip()
                            redacted_text = self.smart_redact_text(original_text)
                            if redacted_text != original_text:
                                # 记录历史
                                operation_history.append({
                                    'row': row,
                                    'col': col,
                                    'original_text': original_text,
                                    'redacted_text': redacted_text,
                                    'original_background': item.background(),
                                    'original_tooltip': item.toolTip()
                                })
                                
                                item.setText(redacted_text)
                                # 标记脱敏的单元格
                                item.setBackground(QColor(255, 235, 235))  # 浅红色背景
                                item.setToolTip(f"已脱敏 - 原文本: {original_text}")
                                redacted_count += 1
            
            # 记录撤销历史
            if operation_history:
                self.excel_redaction_history.append({
                    'type': 'column_range_redaction',
                    'operations': operation_history
                })
            
            if redacted_count > 0:
                QMessageBox.information(self, "脱敏完成", f"已成功脱敏 {redacted_count} 个单元格")
            else:
                QMessageBox.information(self, "提示", "没有找到需要脱敏的内容")

    def mark_table_redaction_all(self):
        """Excel表格全局查找替换脱敏（类似文本和Word的功能）"""
        # 获取当前选中的单元格内容作为查找目标
        selected_items = self.table_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先选择包含要脱敏内容的单元格")
            return
        
        # 使用第一个选中单元格的内容作为查找目标
        target_text = selected_items[0].text().strip()
        if not target_text:
            QMessageBox.warning(self, "警告", "选中的单元格内容为空")
            return
        
        # 询问用户确认
        reply = QMessageBox.question(self, "确认全表脱敏", 
                                   f"确定要对表格中所有包含「{target_text}」的单元格进行脱敏吗？\n请慎重操作！",
                                   QMessageBox.StandardButton.Yes, QMessageBox.StandardButton.No)
        
        if reply != QMessageBox.StandardButton.Yes:
            return
        
        # 记录脱敏的单元格数量和历史记录
        redacted_count = 0
        operation_history = []
        
        # 遍历整个表格查找匹配的内容
        for row in range(self.table_widget.rowCount()):
            for col in range(self.table_widget.columnCount()):
                item = self.table_widget.item(row, col)
                if item and target_text in item.text():
                    original_text = item.text()
                    # 使用智能脱敏替换目标文本
                    redacted_target = self.smart_redact_text(target_text)
                    redacted_text = original_text.replace(target_text, redacted_target)
                    
                    if redacted_text != original_text:
                        # 记录历史
                        operation_history.append({
                            'row': row,
                            'col': col,
                            'original_text': original_text,
                            'redacted_text': redacted_text,
                            'original_background': item.background(),
                            'original_tooltip': item.toolTip()
                        })
                        
                        item.setText(redacted_text)
                        # 标记脱敏的单元格
                        item.setBackground(QColor(255, 235, 235))  # 浅红色背景
                        item.setToolTip(f"已脱敏 - 原文本: {original_text}")
                        redacted_count += 1
        
        # 记录撤销历史
        if operation_history:
            self.excel_redaction_history.append({
                'type': 'table_find_replace_redaction',
                'operations': operation_history,
                'target_text': target_text
            })
        
        if redacted_count > 0:
            QMessageBox.information(self, "脱敏完成", f"已成功脱敏表格中包含「{target_text}」的 {redacted_count} 个单元格")
        else:
            QMessageBox.information(self, "提示", f"表格中没有找到包含「{target_text}」的其他单元格")

    def setup_styles(self):
        # 保持原有配色方案
        palette = self.palette()
        gradient = QLinearGradient(0, 0, 0, self.height())
        gradient.setColorAt(0, QColor(230, 245, 255))
        gradient.setColorAt(1, QColor(180, 220, 255))
        palette.setBrush(QPalette.Window, QBrush(gradient))
        self.setPalette(palette)

        # 统一控件样式
        self.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #3498db;
                border-radius: 5px;
                margin-top: 1ex;
                background-color: rgba(255, 255, 255, 180);
            }
            QGroupBox::title {
                color: #2980b9;
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QTextEdit, QLabel {
                font-size: 11pt;
            }
        """)

    def undo_redaction(self):
        """撤销Excel表格的脱敏操作（基于历史记录的逐步撤销）"""
        if not self.excel_redaction_history:
            QMessageBox.information(self, "提示", "没有可撤销的操作")
            return
        
        # 获取最后一次操作
        last_operation = self.excel_redaction_history.pop()
        
        # 根据操作类型显示确认信息
        op_type = last_operation['type']
        op_operations = last_operation['operations']
        
        type_names = {
            'cell_redaction': '单元格脱敏',
            'row_redaction': '行脱敏',
            'column_redaction': '列脱敏'
        }
        
        op_name = type_names.get(op_type, '脱敏操作')
        
        # 获取用户确认
        reply = QMessageBox.question(self, "确认撤销", 
                                   f"确定要撤销最后一次{op_name}操作吗？\n将影响 {len(op_operations)} 个单元格",
                                   QMessageBox.StandardButton.Yes, QMessageBox.StandardButton.No)
        
        if reply != QMessageBox.StandardButton.Yes:
            # 用户取消，重新添加到历史记录
            self.excel_redaction_history.append(last_operation)
            return
        
        # 执行撤销
        restored_count = 0
        for operation in op_operations:
            item = self.table_widget.item(operation['row'], operation['col'])
            if item:
                # 恢复原始文本
                item.setText(operation['original_text'])
                # 恢复原始背景色
                item.setBackground(operation['original_background'])
                # 恢复原始工具提示
                item.setToolTip(operation['original_tooltip'])
                restored_count += 1
        
        if restored_count > 0:
            QMessageBox.information(self, "撤销完成", f"已成功撤销{op_name}，恢复了 {restored_count} 个单元格")
        else:
            QMessageBox.information(self, "提示", "撤销操作未能恢复任何单元格")
    
    def undo_text_redaction(self):
        """撤销文本编辑器的脱敏操作"""
        if not self.text_redaction_history:
            QMessageBox.information(self, "提示", "没有可撤销的脱敏操作")
            return
        
        # 获取用户确认
        reply = QMessageBox.question(self, "确认撤销", 
                                   f"确定要撤销最后 {len(self.text_redaction_history)} 个脱敏操作吗？",
                                   QMessageBox.StandardButton.Yes, QMessageBox.StandardButton.No)
        
        if reply != QMessageBox.StandardButton.Yes:
            return
        
        # 逆序撤销操作（后进先出）
        restored_count = 0
        while self.text_redaction_history:
            operation = self.text_redaction_history.pop()
            
            if operation.get('type') == 'replace_all':
                # 批量替换的撤销
                self.text_edit.setPlainText(operation['full_original'])
                restored_count += operation['count']
            else:
                # 单个选择替换的撤销
                cursor = self.text_edit.textCursor()
                cursor.setPosition(operation['start'])
                cursor.setPosition(operation['end'], cursor.KeepAnchor)
                cursor.insertText(operation['original'])
                restored_count += 1
        
        if restored_count > 0:
            QMessageBox.information(self, "撤销完成", f"已成功撤销 {restored_count} 处脱敏内容")
    
    def undo_word_redaction(self):
        """撤销Word文档编辑器的脱敏操作"""
        if not self.word_redaction_history:
            QMessageBox.information(self, "提示", "没有可撤销的脱敏操作")
            return
        
        # 获取用户确认
        reply = QMessageBox.question(self, "确认撤销", 
                                   f"确定要撤销最后 {len(self.word_redaction_history)} 个脱敏操作吗？",
                                   QMessageBox.StandardButton.Yes, QMessageBox.StandardButton.No)
        
        if reply != QMessageBox.StandardButton.Yes:
            return
        
        # 逆序撤销操作（后进先出）
        restored_count = 0
        while self.word_redaction_history:
            operation = self.word_redaction_history.pop()
            
            if operation.get('type') == 'replace_all':
                # 批量替换的撤销
                self.word_edit.setPlainText(operation['full_original'])
                restored_count += operation['count']
            else:
                # 单个选择替换的撤销
                cursor = self.word_edit.textCursor()
                cursor.setPosition(operation['start'])
                cursor.setPosition(operation['end'], cursor.KeepAnchor)
                cursor.insertText(operation['original'])
                restored_count += 1
        
        if restored_count > 0:
            QMessageBox.information(self, "撤销完成", f"已成功撤销 {restored_count} 处脱敏内容")

    # 区域撤销功能已移除，仅保留单步撤销

    def undo_current_excel_redaction(self):
        """撤销当前选中区域的Excel脱敏"""
        # 获取当前选中的单元格或区域
        selected_items = self.table_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先选中需要撤销脱敏的单元格")
            return
        
        # 如果只选中一个单元格，进行单元格撤销
        if len(selected_items) == 1:
            self.undo_single_cell_redaction(selected_items[0])
        else:
            # 选中多个单元格，进行区域撤销
            self.undo_region_redaction(selected_items)
    
    def undo_single_cell_redaction(self, cell_item):
        """撤销单个单元格的脱敏"""
        row = cell_item.row()
        col = cell_item.column()
        current_text = cell_item.text()
        
        # 在历史记录中查找该单元格的最新脱敏记录
        restored_count = 0
        for entry in reversed(self.excel_redaction_history):
            # 只处理新格式
            if 'operations' in entry:
                for operation in reversed(entry['operations']):
                    if (operation.get('row') == row and 
                        operation.get('col') == col and 
                        operation.get('redacted_text') == current_text):
                        cell_item.setText(operation['original_text'])
                        cell_item.setBackground(operation.get('original_background', QColor()))
                        cell_item.setToolTip(operation.get('original_tooltip', ''))
                        entry['operations'].remove(operation)
                        if not entry['operations']:
                            self.excel_redaction_history.remove(entry)
                        restored_count = 1
                        break
                if restored_count > 0:
                    break
        
        if restored_count > 0:
            QMessageBox.information(self, "撤销完成", f"已撤销单元格({row+1}, {col+1})的脱敏内容")
        else:
            QMessageBox.information(self, "提示", f"未找到单元格({row+1}, {col+1})的脱敏历史记录")
    
    def undo_region_redaction(self, selected_items):
        """撤销选中区域的脱敏"""
        # 获取选中单元格的位置集合
        selected_positions = {(item.row(), item.column()) for item in selected_items}
        
        restored_count = 0
        operations_to_remove = []  # 记录需要移除的操作
        entries_to_remove = []     # 记录需要移除的整个entry
        
        # 遍历历史记录
        for entry in reversed(self.excel_redaction_history):
            if 'operations' in entry:
                for operation in reversed(entry['operations']):
                    pos = (operation.get('row'), operation.get('col'))
                    if pos in selected_positions:
                        cell_item = self.table_widget.item(pos[0], pos[1])
                        if (cell_item and cell_item.text() == operation.get('redacted_text')):
                            cell_item.setText(operation['original_text'])
                            cell_item.setBackground(operation.get('original_background', QColor()))
                            cell_item.setToolTip(operation.get('original_tooltip', ''))
                            operations_to_remove.append((entry, operation))
                            restored_count += 1
        for entry, operation in operations_to_remove:
            if operation in entry['operations']:
                entry['operations'].remove(operation)
                if not entry['operations']:
                    entries_to_remove.append(entry)
        for entry in entries_to_remove:
            if entry in self.excel_redaction_history:
                self.excel_redaction_history.remove(entry)
        if restored_count > 0:
            QMessageBox.information(self, "撤销完成", f"已撤销选中区域内 {restored_count} 个单元格的脱敏内容")
        else:
            QMessageBox.information(self, "提示", "未找到选中区域内的脱敏历史记录")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = UniversalRedactionTool()
    window.show()
    sys.exit(app.exec_())

"""
OpenDataLoader PDF 集成模块
用于岩土工程报告的高精度 PDF 数据提取
特别优化表格数据提取（水位、埋深、腐蚀性等）
"""

import json
from pathlib import Path
from typing import List, Dict, Any, Optional

try:
    import opendataloader_pdf
    OPEN_DATALOADER_AVAILABLE = True
except ImportError:
    OPEN_DATALOADER_AVAILABLE = False
    print("警告：opendataloader-pdf 未安装")
    print("安装方式：pip install -U opendataloader-pdf")
    print("注意：需要 Java 11+ 在系统 PATH 中")

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False


class OpenDataLoaderReader:
    """使用 OpenDataLoader PDF 的读取器（推荐）"""

    def __init__(self, pdf_path: Path):
        self.pdf_path = Path(pdf_path)
        self.json_data = None
        self.text = ""
        self.tables = []

    def extract_all(self, output_format: str = "json") -> bool:
        """
        使用 OpenDataLoader 提取所有内容

        Args:
            output_format: 输出格式，支持 "json", "markdown", "html"

        Returns:
            bool: 是否成功
        """
        if not OPEN_DATALOADER_AVAILABLE:
            print("OpenDataLoader 不可用")
            return False

        try:
            # 临时输出目录
            temp_output = self.pdf_path.parent / ".temp_opendataloader"
            temp_output.mkdir(exist_ok=True)

            # 调用 OpenDataLoader 转换
            opendataloader_pdf.convert(
                input_path=[str(self.pdf_path)],
                output_dir=str(temp_output),
                format=output_format,
                table_method="auto",  # 自动表格检测
                reading_order="xycut",  # 处理多栏文档
            )

            # 读取 JSON 结果
            json_file = temp_output / f"{self.pdf_path.stem}.json"
            if json_file.exists():
                with open(json_file, 'r', encoding='utf-8') as f:
                    self.json_data = json.load(f)
                self._parse_json_data()
                return True

            return False

        except Exception as e:
            print(f"OpenDataLoader 提取失败: {e}")
            return False

    def _parse_json_data(self):
        """解析 OpenDataLoader 输出的 JSON 数据"""
        if not self.json_data:
            return

        # 提取文本
        if "pages" in self.json_data:
            for page in self.json_data["pages"]:
                if "elements" in page:
                    for element in page["elements"]:
                        if element.get("type") == "text":
                            self.text += element.get("content", "") + "\n"

        # 提取表格
        if "pages" in self.json_data:
            for page_num, page in enumerate(self.json_data["pages"], 1):
                if "elements" in page:
                    for element in page["elements"]:
                        if element.get("type") == "table":
                            table_data = self._extract_table_data(element)
                            self.tables.append({
                                "page": page_num,
                                "data": table_data,
                                "bbox": element.get("bbox", {})
                            })

    def _extract_table_data(self, table_element: Dict) -> List[List[str]]:
        """从表格元素中提取数据"""
        table_data = []

        if "cells" in table_element:
            # OpenDataLoader v2.0 格式
            for row in table_element["cells"]:
                row_data = [cell.get("text", "").strip() for cell in row]
                if any(row_data):  # 跳过空行
                    table_data.append(row_data)
        elif "rows" in table_element:
            # 备选格式
            for row in table_element["rows"]:
                row_data = [cell.get("content", "").strip() for cell in row.get("cells", [])]
                if any(row_data):
                    table_data.append(row_data)

        return table_data

    def extract_text(self) -> str:
        """提取文本"""
        return self.text

    def extract_tables(self) -> List[Dict]:
        """提取表格（带页码和位置信息）"""
        return self.tables

    def get_table_by_keywords(self, keywords: List[str]) -> Optional[List[List[str]]]:
        """
        根据关键词查找表格

        Args:
            keywords: 关键词列表，匹配任一关键词即可

        Returns:
            匹配的表格数据，或 None
        """
        for table in self.tables:
            table_str = json.dumps(table["data"], ensure_ascii=False)
            if any(kw in table_str for kw in keywords):
                return table["data"]
        return None

    def get_water_level_data(self) -> Dict[str, Any]:
        """提取水位数据"""
        water_keywords = ["水位", "地下水", "埋深", "water"]
        table = self.get_table_by_keywords(water_keywords)

        result = {"found": False, "data": []}
        if table:
            result["found"] = True
            result["data"] = table
        return result

    def get_corrosion_data(self) -> Dict[str, Any]:
        """提取腐蚀性数据"""
        corrosion_keywords = ["腐蚀", "腐蚀性", "SO4", "Cl", "易溶盐", "corrosion"]
        table = self.get_table_by_keywords(corrosion_keywords)

        result = {"found": False, "data": []}
        if table:
            result["found"] = True
            result["data"] = table
        return result

    def get_spt_data(self) -> Dict[str, Any]:
        """提取标贯数据"""
        spt_keywords = ["标贯", "SPT", "标准贯入", "N63.5"]
        table = self.get_table_by_keywords(spt_keywords)

        result = {"found": False, "data": []}
        if table:
            result["found"] = True
            result["data"] = table
        return result


class PDFPlumberReader:
    """使用 pdfplumber 的读取器（备选）"""

    def __init__(self, pdf_path: Path):
        self.pdf_path = Path(pdf_path)

    def extract_text(self) -> str:
        """提取文本"""
        if not PDFPLUMBER_AVAILABLE:
            return ""

        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                text = "\n".join([page.extract_text() or "" for page in pdf.pages])
            return text
        except Exception as e:
            print(f"pdfplumber 读取失败: {e}")
            return ""

    def extract_tables(self) -> List[Dict]:
        """提取表格"""
        if not PDFPLUMBER_AVAILABLE:
            return []

        tables = []
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    page_tables = page.extract_tables()
                    if page_tables:
                        for table_data in page_tables:
                            if table_data:
                                tables.append({
                                    "page": page_num,
                                    "data": table_data,
                                    "bbox": {}
                                })
        except Exception as e:
            print(f"pdfplumber 提取表格失败: {e}")

        return tables


def get_best_reader(pdf_path: Path, prefer_opendataloader: bool = True) -> object:
    """
    获取最佳可用的 PDF 读取器

    Args:
        pdf_path: PDF 文件路径
        prefer_opendataloader: 是否优先使用 OpenDataLoader

    Returns:
        PDF 读取器实例
    """
    if prefer_opendataloader and OPEN_DATALOADER_AVAILABLE:
        return OpenDataLoaderReader(pdf_path)
    elif PDFPLUMBER_AVAILABLE:
        return PDFPlumberReader(pdf_path)
    else:
        raise ImportError("无可用的 PDF 读取器，请安装 opendataloader-pdf 或 pdfplumber")


# ============================================================
# 岩土工程专用提取函数
# ============================================================

def extract_geotech_data(pdf_path: Path) -> Dict[str, Any]:
    """
    从岩土工程 PDF 中提取所有关键数据

    Args:
        pdf_path: PDF 文件路径

    Returns:
        包含所有提取数据的字典
    """
    reader = get_best_reader(pdf_path)

    result = {
        "pdf_name": pdf_path.name,
        "text": "",
        "tables": [],
        "water_level": {"found": False, "data": []},
        "corrosion": {"found": False, "data": []},
        "spt": {"found": False, "data": []},
        "wave_velocity": {"found": False, "data": []},
    }

    # 如果是 OpenDataLoaderReader，需要先执行 extract_all 再取数据
    if isinstance(reader, OpenDataLoaderReader):
        if not reader.json_data:
            reader.extract_all()

    # 提取文本
    result["text"] = reader.extract_text()

    # 提取表格
    result["tables"] = reader.extract_tables()

    # 如果是 OpenDataLoaderReader，使用专用方法提取业务数据
    if isinstance(reader, OpenDataLoaderReader):
        result["water_level"] = reader.get_water_level_data()
        result["corrosion"] = reader.get_corrosion_data()
        result["spt"] = reader.get_spt_data()

        # 波速数据
        wave_keywords = ["波速", "剪切波速", "Vs", "wave", "velocity"]
        wave_table = reader.get_table_by_keywords(wave_keywords)
        if wave_table:
            result["wave_velocity"] = {"found": True, "data": wave_table}
    else:
        # pdfplumber 备选：从表格中搜索
        keywords_map = {
            "water_level": ["水位", "地下水", "埋深"],
            "corrosion": ["腐蚀", "腐蚀性", "SO4", "Cl"],
            "spt": ["标贯", "SPT", "标准贯入"],
            "wave_velocity": ["波速", "剪切波速", "Vs"],
        }

        for key, keywords in keywords_map.items():
            for table in result["tables"]:
                table_str = json.dumps(table["data"], ensure_ascii=False)
                if any(kw in table_str for kw in keywords):
                    result[key] = {"found": True, "data": table["data"]}
                    break

    return result


def extract_batch_pdfs(pdf_folder: Path, pattern: str = "*.pdf") -> Dict[str, Any]:
    """
    批量提取文件夹中的 PDF 数据

    Args:
        pdf_folder: PDF 文件夹
        pattern: 文件匹配模式

    Returns:
        所有 PDF 的提取结果
    """
    pdf_folder = Path(pdf_folder)
    results = {}

    pdfs = list(pdf_folder.glob(pattern))
    pdfs.extend(pdf_folder.glob(f"**/{pattern}"))
    pdfs = list(dict.fromkeys(pdfs))

    print(f"发现 {len(pdfs)} 个 PDF 文件")

    for pdf_path in sorted(pdfs):
        print(f"  处理: {pdf_path.name}")
        results[pdf_path.name] = extract_geotech_data(pdf_path)

    return results


# ============================================================
# 测试代码
# ============================================================

if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        # 测试指定文件
        test_file = Path(sys.argv[1])
        data = extract_geotech_data(test_file)

        print(f"\n文件: {data['pdf_name']}")
        print(f"文本长度: {len(data['text'])} 字符")
        print(f"表格数量: {len(data['tables'])}")
        print(f"水位数据: {'找到' if data['water_level']['found'] else '未找到'}")
        print(f"腐蚀性数据: {'找到' if data['corrosion']['found'] else '未找到'}")
        print(f"标贯数据: {'找到' if data['spt']['found'] else '未找到'}")
        print(f"波速数据: {'找到' if data['wave_velocity']['found'] else '未找到'}")

        # 显示找到的数据
        if data['water_level']['found']:
            print("\n=== 水位数据 ===")
            for row in data['water_level']['data'][:5]:
                print(row)

        if data['corrosion']['found']:
            print("\n=== 腐蚀性数据 ===")
            for row in data['corrosion']['data'][:5]:
                print(row)
    else:
        # 显示可用性
        print("PDF 读取器状态:")
        print(f"  OpenDataLoader PDF: {'✅ 可用' if OPEN_DATALOADER_AVAILABLE else '❌ 不可用'}")
        print(f"  pdfplumber: {'✅ 可用' if PDFPLUMBER_AVAILABLE else '❌ 不可用'}")
        print("\n使用方式:")
        print("  python opendataloader_pdf_reader.py <pdf文件路径>")
        print("  python opendataloader_pdf_reader.py <pdf文件夹路径>")

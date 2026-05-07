import pandas as pd
import re
from pathlib import Path
from typing import Dict, List, Any
import warnings
warnings.filterwarnings('ignore')

class ExcelQuotationExtractor:
    """Excel报价表提取器"""
    
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.original_data = {}
        self.results = []
        
    def load_excel(self):
        """加载所有sheet"""
        excel_file = pd.ExcelFile(self.file_path)
        for sheet_name in excel_file.sheet_names:
            # 跳过系统表
            if sheet_name in ['等级报价', '其他特殊价格', '商派E F等级邮编', '审批权限流程和规则']:
                continue
            self.original_data[sheet_name] = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None)
            
    def detect_price_tables(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        """检测并提取报价表"""
        tables = []
        i = 0
        rows = df.values.tolist()
        
        while i < len(rows):
            row = rows[i]
            row_str = ' '.join([str(cell) if pd.notna(cell) else '' for cell in row])
            
            # 检测表头行
            if '重量段' in row_str and 'KG' in row_str and ('国家' in row_str or '运费' in row_str):
                product_code = self._extract_product_code(rows, i-2) if i >= 2 else ''
                update_time = self._extract_update_time(rows, i-1) if i >= 1 else ''
                
                # 获取列索引
                col_indices = self._get_column_indices(row)
                
                # 查找数据行
                data_rows = []
                j = i + 1
                while j < len(rows):
                    data_row = rows[j]
                    # 检查是否为空行或新表头
                    if self._is_empty_row(data_row):
                        j += 1
                        continue
                    if self._is_new_table_start(data_row):
                        break
                    
                    # 提取数据
                    record = self._extract_record(data_row, col_indices)
                    if record and record.get('重量段'):
                        record['产品代码'] = product_code
                        record['更新时间'] = update_time
                        record['sheet名称'] = sheet_name
                        data_rows.append(record)
                    j += 1
                
                if data_rows:
                    tables.append({
                        '产品代码': product_code,
                        '更新时间': update_time,
                        '数据': data_rows,
                        'sheet名称': sheet_name
                    })
                i = j
            else:
                i += 1
        return tables
    
    def _extract_product_code(self, rows: List[List], idx: int) -> str:
        """提取产品代码"""
        if idx < 0 or idx >= len(rows):
            return ''
        row = rows[idx]
        # 产品代码通常在第一个非空单元格
        for cell in row:
            if pd.notna(cell) and isinstance(cell, str) and len(cell) > 2:
                # 清理产品代码
                code = cell.strip()
                # 常见模式：XXX-XXX, XXX/XXX等
                if re.search(r'[A-Z]{2,}[/\-]?[A-Z0-9]+', code):
                    return code
                if '(' in code or '）' in code:
                    # 提取括号前的内容
                    match = re.match(r'^([^\(（]+)', code)
                    if match:
                        return match.group(1).strip()
        return ''
    
    def _extract_update_time(self, rows: List[List], idx: int) -> str:
        """提取更新时间"""
        if idx < 0 or idx >= len(rows):
            return ''
        row = rows[idx]
        for cell in row:
            if pd.notna(cell) and isinstance(cell, str):
                # 匹配日期格式
                date_patterns = [
                    r'(\d{4}年\d{1,2}月\d{1,2}日)',
                    r'(\d{4}-\d{1,2}-\d{1,2})',
                    r'(\d{4}/\d{1,2}/\d{1,2})',
                    r'(\d{1,2}月\d{1,2}日)'
                ]
                for pattern in date_patterns:
                    match = re.search(pattern, cell)
                    if match:
                        return match.group(1)
        return ''
    
    def _get_column_indices(self, header_row: List) -> Dict:
        """获取列索引映射"""
        indices = {
            '国家': -1, '重量段': -1, '运费': -1, '处理费': -1,
            '上网参考时效': -1, '尺寸限制': -1
        }
        
        for idx, cell in enumerate(header_row):
            cell_str = str(cell).strip() if pd.notna(cell) else ''
            for key in indices.keys():
                if key in cell_str and indices[key] == -1:
                    indices[key] = idx
        return indices
    
    def _extract_record(self, row: List, col_indices: Dict) -> Dict:
        """提取单行记录"""
        record = {}
        
        # 提取国家
        if col_indices['国家'] >= 0 and col_indices['国家'] < len(row):
            country = row[col_indices['国家']]
            if pd.notna(country):
                record['国家'] = str(country).strip()
        
        # 提取重量段
        if col_indices['重量段'] >= 0 and col_indices['重量段'] < len(row):
            weight = row[col_indices['重量段']]
            if pd.notna(weight):
                record['重量段'] = str(weight).strip()
        
        # 提取运费
        if col_indices['运费'] >= 0 and col_indices['运费'] < len(row):
            shipping = row[col_indices['运费']]
            if pd.notna(shipping):
                try:
                    record['运费(元/KG)'] = float(shipping)
                except:
                    record['运费(元/KG)'] = shipping
        
        # 提取处理费
        if col_indices['处理费'] >= 0 and col_indices['处理费'] < len(row):
            fee = row[col_indices['处理费']]
            if pd.notna(fee):
                try:
                    record['处理费(元/票)'] = float(fee)
                except:
                    record['处理费(元/票)'] = fee
        
        return record if record.get('重量段') else None
    
    def _is_empty_row(self, row: List) -> bool:
        """检查是否为空行"""
        return all(pd.isna(cell) or str(cell).strip() == '' for cell in row)
    
    def _is_new_table_start(self, row: List) -> bool:
        """检查是否是新表格开始"""
        row_str = ' '.join([str(cell) for cell in row if pd.notna(cell)])
        return '重量段' in row_str and 'KG' in row_str
    
    def extract_all(self) -> List[Dict]:
        """提取所有报价表"""
        self.load_excel()
        all_tables = []
        
        for sheet_name, df in self.original_data.items():
            print(f"正在处理: {sheet_name}")
            tables = self.detect_price_tables(df, sheet_name)
            all_tables.extend(tables)
            
        self.results = all_tables
        return all_tables
    
    def to_dataframe(self) -> pd.DataFrame:
        """转换为DataFrame"""
        if not self.results:
            self.extract_all()
        
        all_records = []
        for table in self.results:
            for record in table['数据']:
                all_records.append(record)
        
        return pd.DataFrame(all_records)
    
    def save_to_excel(self, output_path: str):
        """保存结果到Excel"""
        df = self.to_dataframe()
        df.to_excel(output_path, index=False)
        print(f"结果已保存到: {output_path}")
        return df
    
    def print_summary(self):
        """打印摘要"""
        if not self.results:
            self.extract_all()
        
        print(f"\n{'='*60}")
        print(f"共发现 {len(self.results)} 个报价表")
        print(f"{'='*60}")
        
        for table in self.results:
            print(f"\n产品代码: {table['产品代码']}")
            print(f"更新时间: {table['更新时间']}")
            print(f"Sheet: {table['sheet名称']}")
            print(f"数据行数: {len(table['数据'])}")
            print("-" * 40)


# 使用AI模型辅助识别（使用OpenAI API）
class AIEnhancedExtractor(ExcelQuotationExtractor):
    """AI增强版提取器"""
    
    def __init__(self, file_path: str, api_key: str = None):
        super().__init__(file_path)
        self.api_key = api_key
    
    def analyze_with_ai(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        """使用AI模型分析表格结构"""
        try:
            import openai
            
            openai.api_key = self.api_key
            
            # 将Excel片段转为文本
            sample_text = df.head(50).to_string()
            
            prompt = f"""
            分析以下Excel表格内容，这是一个物流报价表。
            请识别出：
            1. 所有的产品代码
            2. 每个产品的更新时间
            3. 每个产品的价格表（国家、重量段、运费、处理费）
            
            表格内容：
            {sample_text}
            
            请以JSON格式返回结果。
            """
            
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3
            )
            
            # 解析AI返回的结果
            import json
            result = json.loads(response.choices[0].message.content)
            return result.get('tables', [])
            
        except ImportError:
            print("未安装openai库，请先安装: pip install openai")
            return []
        except Exception as e:
            print(f"AI分析失败: {e}")
            return []


# 使用LangChain框架
class LangChainExtractor:
    """使用LangChain进行表格提取"""
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        
    def extract_with_langchain(self):
        """使用LangChain + LLM提取"""
        try:
            from langchain.document_loaders import UnstructuredExcelLoader
            from langchain.text_splitter import CharacterTextSplitter
            from langchain.chains import create_extraction_chain
            from langchain.llms import OpenAI
            
            # 加载Excel
            loader = UnstructuredExcelLoader(self.file_path, mode="elements")
            documents = loader.load()
            
            # 定义提取schema
            schema = {
                "properties": {
                    "product_code": {"type": "string"},
                    "update_time": {"type": "string"},
                    "country": {"type": "string"},
                    "weight_range": {"type": "string"},
                    "shipping_fee": {"type": "number"},
                    "handling_fee": {"type": "number"}
                },
                "required": ["product_code", "country", "weight_range"]
            }
            
            llm = OpenAI(temperature=0)
            chain = create_extraction_chain(llm, schema)
            
            # 提取信息
            result = chain.run(documents)
            return result
            
        except ImportError:
            print("请安装langchain: pip install langchain")
            return None


# 使用示例
def main():
    # 初始化提取器
    extractor = ExcelQuotationExtractor('出口易物流推广报价表2026.04.22.xlsx')
    
    # 提取所有报价表
    tables = extractor.extract_all()
    
    # 打印摘要
    extractor.print_summary()
    
    # 转换为DataFrame并保存
    df = extractor.save_to_excel('提取的报价数据.xlsx')
    
    # 查看结果
    print(f"\n提取的数据预览:")
    print(df.head(20))
    
    # 按产品代码分组统计
    if len(df) > 0:
        print(f"\n按产品代码统计:")
        print(df.groupby('产品代码').size())
        
        # 保存完整结果
        extractor.save_to_excel('报价数据_完整.xlsx')
    
    return df


# 高级：使用OpenAI函数调用
class OpenAIFunctionExtractor:
    """使用OpenAI函数调用能力"""
    
    def __init__(self, api_key: str):
        import openai
        openai.api_key = api_key
        self.client = openai
        
    def extract_price_table(self, excel_content: str) -> Dict:
        """使用函数调用提取价格表"""
        
        functions = [
            {
                "name": "extract_price_table",
                "description": "从Excel文本中提取物流报价表信息",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "tables": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "product_code": {"type": "string"},
                                    "update_time": {"type": "string"},
                                    "rates": {
                                        "type": "array",
                                        "items": {
                                            "type": "object",
                                            "properties": {
                                                "country": {"type": "string"},
                                                "weight_range": {"type": "string"},
                                                "shipping_fee": {"type": "number"},
                                                "handling_fee": {"type": "number"}
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        ]
        
        response = self.client.ChatCompletion.create(
            model="gpt-3.5-turbo-16k",
            messages=[
                {"role": "system", "content": "你是一个Excel表格分析专家，请提取物流报价表信息。"},
                {"role": "user", "content": excel_content[:15000]}  # 限制长度
            ],
            functions=functions,
            function_call={"name": "extract_price_table"},
            temperature=0
        )
        
        import json
        result = json.loads(response.choices[0].message.function_call.arguments)
        return result


if __name__ == "__main__":
    # 运行主程序
    df = main()
    
    # 如果需要AI增强版（需要OpenAI API Key）
    # ai_extractor = AIEnhancedExtractor('出口易物流推广报价表2026.04.22.xlsx', 'your-api-key')
    # tables = ai_extractor.extract_all()
import re
import os
from collections import defaultdict, OrderedDict
from datetime import datetime

class EnhancedSASAnalyzer:
    def __init__(self):
        self.reset()
    
    def reset(self):
        """Reset analyzer state for new analysis"""
        self.analysis_results = {
            'function_blocks': [],
            'datasets_created': defaultdict(list),
            'datasets_used': defaultdict(list),
            'procedures_used': defaultdict(list),
            'macros_defined': defaultdict(dict),
            'macros_called': defaultdict(list),
            'libraries_defined': defaultdict(list),
            'sql_statements': defaultdict(list),
            'variables_used': defaultdict(list),
            'file_operations': defaultdict(list),
            'control_structures': defaultdict(list),
            'include_files': defaultdict(list),
            'system_functions': defaultdict(list),
            'call_routines': defaultdict(list),
            'formats': defaultdict(list),
            'hash_objects': defaultdict(list),
            'ods_statements': defaultdict(list),
            'timeframe_start': None,
            'timeframe_end': None,
            'code_complexity': {},
            'line_analysis': {}
        }
        self.current_blocks = []
        self.include_stack = []
    
    def analyze_file(self, file_path):
        """Analyze SAS code from external text file"""
        if not os.path.exists(file_path):
            print(f"❌ Error: File '{file_path}' not found.")
            return None
            
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                lines = file.readlines()
            print(f"✅ Successfully loaded file: {file_path} ({len(lines)} lines)")
            return self.analyze_lines(lines, file_path)
        except Exception as e:
            print(f"❌ Error reading file {file_path}: {e}")
            return None
    
    def analyze_lines(self, lines, source_file=None):
        """Main analysis function for list of lines"""
        self.reset()
        total_lines = len(lines)
        
        for line_num, line in enumerate(lines, 1):
            self._analyze_single_line(line, line_num, total_lines, source_file)
        
        self._finalize_open_blocks(total_lines)
        self._calculate_complexity_metrics()
        
        return self.analysis_results
    
    def _analyze_single_line(self, line, line_num, total_lines, source_file):
        """Analyze a single line of SAS code"""
        original_line = line.strip()
        cleaned_line = self._clean_line(line).strip()
        
        if not cleaned_line:
            return
        
        self.analysis_results['line_analysis'][line_num] = {
            'original': original_line,
            'cleaned': cleaned_line,
            'type': self._classify_line(cleaned_line),
            'source_file': source_file
        }
        
        # Check for block starts
        self._check_block_start(cleaned_line, line_num)
        
        # Analyze all components
        self._analyze_data_operations(cleaned_line, line_num)
        self._analyze_procedures(cleaned_line, line_num)
        self._analyze_macros(cleaned_line, line_num)
        self._analyze_sql_operations(cleaned_line, line_num)
        self._analyze_control_structures(cleaned_line, line_num)
        self._analyze_file_operations(cleaned_line, line_num)
        self._analyze_variables(cleaned_line, line_num)
        self._analyze_include_files(cleaned_line, line_num)
        self._analyze_system_functions(cleaned_line, line_num)
        self._analyze_call_routines(cleaned_line, line_num)
        self._analyze_formats(cleaned_line, line_num)
        self._analyze_hash_objects(cleaned_line, line_num)
        self._analyze_ods_statements(cleaned_line, line_num)
        self._analyze_timeframes(original_line, line_num)
        
        # Check for block ends
        self._check_block_end(cleaned_line, line_num)
    
    def _clean_line(self, line):
        """Clean line by removing comments"""
        line = re.sub(r'/\*.*?\*/', '', line)
        line = re.sub(r'^\s*\*.*$', '', line)
        return line.upper()
    
    def _classify_line(self, line):
        """Classify the type of SAS statement"""
        if re.match(r'^\s*%INCLUDE\s+', line):
            return 'INCLUDE'
        elif re.match(r'^\s*DATA\s+', line):
            return 'DATA_STEP'
        elif re.match(r'^\s*PROC\s+', line):
            return 'PROCEDURE'
        elif re.match(r'^\s*%MACRO\s+', line):
            return 'MACRO_DEF'
        elif re.match(r'^\s*%', line):
            return 'MACRO_CALL'
        elif re.match(r'^\s*LIBNAME\s+', line):
            return 'LIBRARY'
        elif re.match(r'^\s*ODS\s+', line):
            return 'ODS'
        elif 'RUN;' in line or 'QUIT;' in line:
            return 'TERMINATOR'
        else:
            return 'STATEMENT'
    
    def _analyze_include_files(self, line, line_num):
        """Analyze %INCLUDE statements"""
        include_patterns = [
            r'%INCLUDE\s+["\']([^"\']+)["\']',
            r'%INCLUDE\s+([A-Z_][A-Z0-9_]*)',
            r'%INCLUDE\s+([^\s;]+)',
        ]
        
        for pattern in include_patterns:
            matches = re.findall(pattern, line)
            for match in matches:
                self.analysis_results['include_files'][match].append(line_num)
    
    def _analyze_system_functions(self, line, line_num):
        """Analyze SAS system functions"""
        sas_functions = [
            'SUM', 'MEAN', 'MIN', 'MAX', 'COUNT', 'N', 'NMISS',
            'SUBSTR', 'TRIM', 'STRIP', 'LEFT', 'RIGHT', 'LENGTH',
            'UPCASE', 'LOWCASE', 'PROPCASE', 'COMPRESS', 'TRANSLATE',
            'INDEX', 'FIND', 'SCAN', 'CATS', 'CATX', 'CAT',
            'INPUT', 'PUT', 'ROUND', 'CEIL', 'FLOOR', 'INT', 'ABS',
            'LOG', 'EXP', 'SQRT', 'SIN', 'COS', 'TAN',
            'TODAY', 'DATE', 'DATETIME', 'TIME', 'DATEPART', 'TIMEPART',
            'YEAR', 'MONTH', 'DAY', 'WEEKDAY', 'MDY', 'YMD',
            'INTCK', 'INTNX', 'DATDIF', 'JULDATE',
            'COALESCEC', 'COALESCE', 'IFC', 'IFN', 'MISSING',
            'RAND', 'RANUNI', 'NORMAL', 'GAMMA', 'BETA'
        ]
        
        for func in sas_functions:
            pattern = rf'\b{func}\s*\('
            if re.search(pattern, line):
                self.analysis_results['system_functions'][func].append(line_num)
    
    def _analyze_call_routines(self, line, line_num):
        """Analyze CALL routines"""
        call_routines = [
            'SYMPUT', 'SYMPUTX', 'SYMGET', 'SYMGETN',
            'EXECUTE', 'SYSTEM', 'FILENAME', 'LIBNAME',
            'STREAMINIT', 'RANUNI', 'RANTBL', 'VNAME',
            'LABEL', 'MISSING', 'SORTC', 'SORTN'
        ]
        
        for routine in call_routines:
            pattern = rf'\bCALL\s+{routine}\b'
            if re.search(pattern, line):
                self.analysis_results['call_routines'][f'CALL_{routine}'].append(line_num)
    
    def _analyze_formats(self, line, line_num):
        """Analyze FORMAT and INFORMAT statements"""
        if re.search(r'\bPROC\s+FORMAT\b', line):
            self.analysis_results['formats']['PROC_FORMAT'].append(line_num)
        
        if re.search(r'\bVALUE\s+[A-Z_][A-Z0-9_]*', line):
            value_match = re.search(r'VALUE\s+([A-Z_][A-Z0-9_]*)', line)
            if value_match:
                format_name = value_match.group(1)
                self.analysis_results['formats'][f'VALUE_{format_name}'].append(line_num)
        
        if re.search(r'\bINFORMAT\s+', line):
            self.analysis_results['formats']['INFORMAT'].append(line_num)
    
    def _analyze_hash_objects(self, line, line_num):
        """Analyze Hash objects"""
        hash_patterns = [
            ('DECLARE_HASH', r'DECLARE\s+HASH\s+([A-Z_][A-Z0-9_]*)'),
            ('DECLARE_HITER', r'DECLARE\s+HITER\s+([A-Z_][A-Z0-9_]*)'),
            ('DEFINEKEY', r'([A-Z_][A-Z0-9_]*)\.DEFINEKEY'),
            ('DEFINEDATA', r'([A-Z_][A-Z0-9_]*)\.DEFINEDATA'),
            ('DEFINEDONE', r'([A-Z_][A-Z0-9_]*)\.DEFINEDONE'),
            ('ADD', r'([A-Z_][A-Z0-9_]*)\.ADD'),
            ('FIND', r'([A-Z_][A-Z0-9_]*)\.FIND'),
            ('CHECK', r'([A-Z_][A-Z0-9_]*)\.CHECK')
        ]
        
        for operation, pattern in hash_patterns:
            matches = re.findall(pattern, line)
            for match in matches:
                self.analysis_results['hash_objects'][f'HASH_{operation}_{match}'].append(line_num)
    
    def _analyze_ods_statements(self, line, line_num):
        """Analyze ODS statements"""
        ods_operations = [
            'HTML', 'PDF', 'RTF', 'EXCEL', 'POWERPOINT', 'CSV',
            'LISTING', 'OUTPUT', 'TRACE', 'SELECT', 'EXCLUDE',
            'GRAPHICS', 'RESULTS', 'DESTINATIONS'
        ]
        
        for operation in ods_operations:
            if re.search(rf'\bODS\s+{operation}\b', line):
                self.analysis_results['ods_statements'][f'ODS_{operation}'].append(line_num)
    
    def _check_block_start(self, line, line_num):
        """Check if line starts a new block"""
        # DATA step
        data_match = re.match(r'^\s*DATA\s+([A-Z_][A-Z0-9_.]*(?:\s+[A-Z_][A-Z0-9_.]*)*)', line)
        if data_match:
            datasets = data_match.group(1).split()
            block_info = {
                'type': 'DATA',
                'name': f"DATA {' '.join(datasets)}",
                'start_line': line_num,
                'end_line': None,
                'datasets': datasets
            }
            self.current_blocks.append(block_info)
            return
        
        # PROC statements
        proc_match = re.match(r'^\s*PROC\s+([A-Z]+)(?:\s+DATA\s*=\s*([A-Z_][A-Z0-9_.]*))?\s*;?', line)
        if proc_match:
            proc_name = proc_match.group(1)
            dataset = proc_match.group(2) if proc_match.group(2) else 'UNKNOWN'
            block_info = {
                'type': 'PROC',
                'name': f"PROC {proc_name}",
                'start_line': line_num,
                'end_line': None,
                'proc_name': proc_name,
                'dataset': dataset
            }
            self.current_blocks.append(block_info)
            return
        
        # Macro definitions
        macro_match = re.match(r'^\s*%MACRO\s+([A-Z_][A-Z0-9_]*)(\([^)]*\))?\s*;?', line)
        if macro_match:
            macro_name = macro_match.group(1)
            params = macro_match.group(2) if macro_match.group(2) else ''
            block_info = {
                'type': 'MACRO',
                'name': f"%MACRO {macro_name}",
                'start_line': line_num,
                'end_line': None,
                'macro_name': macro_name,
                'parameters': params
            }
            self.current_blocks.append(block_info)
    
    def _check_block_end(self, line, line_num):
        """Check if line ends current block"""
        if 'RUN;' in line or 'QUIT;' in line or '%MEND' in line:
            if self.current_blocks:
                block = self.current_blocks.pop()
                block['end_line'] = line_num
                self.analysis_results['function_blocks'].append(block)
    
    def _finalize_open_blocks(self, total_lines):
        """Close any remaining open blocks"""
        while self.current_blocks:
            block = self.current_blocks.pop()
            block['end_line'] = total_lines
            self.analysis_results['function_blocks'].append(block)
    
    def _analyze_data_operations(self, line, line_num):
        """Analyze DATA step operations"""
        # Dataset creation
        data_match = re.search(r'DATA\s+([A-Z_][A-Z0-9_.]*(?:\s+[A-Z_][A-Z0-9_.]*)*)', line)
        if data_match:
            datasets = data_match.group(1).split()
            for dataset in datasets:
                self.analysis_results['datasets_created'][dataset].append(line_num)
        
        # Dataset usage
        usage_patterns = [
            ('SET', r'SET\s+([A-Z_][A-Z0-9_.]*(?:\s+[A-Z_][A-Z0-9_.]*)*?)(?:\s|;|$)'),
            ('MERGE', r'MERGE\s+([A-Z_][A-Z0-9_.]*(?:\s+[A-Z_][A-Z0-9_.]*)*?)(?:\s|;|$)'),
            ('UPDATE', r'UPDATE\s+([A-Z_][A-Z0-9_.]*(?:\s+[A-Z_][A-Z0-9_.]*)*?)(?:\s|;|$)')
        ]
        
        for operation, pattern in usage_patterns:
            matches = re.findall(pattern, line)
            for match in matches:
                datasets = match.split()
                for dataset in datasets:
                    self.analysis_results['datasets_used'][f"{operation}_{dataset}"].append(line_num)
    
    def _analyze_procedures(self, line, line_num):
        """Analyze PROC statements"""
        proc_match = re.search(r'PROC\s+([A-Z]+)(?:\s+DATA\s*=\s*([A-Z_][A-Z0-9_.]*))?\s*;?', line)
        if proc_match:
            proc_name = proc_match.group(1)
            dataset = proc_match.group(2) if proc_match.group(2) else 'UNKNOWN'
            self.analysis_results['procedures_used'][proc_name].append({
                'line': line_num,
                'dataset': dataset
            })
    
    def _analyze_macros(self, line, line_num):
        """Analyze macro definitions and calls"""
        # Macro definitions
        macro_def_match = re.search(r'%MACRO\s+([A-Z_][A-Z0-9_]*)(\([^)]*\))?\s*;?', line)
        if macro_def_match:
            macro_name = macro_def_match.group(1)
            params = macro_def_match.group(2) if macro_def_match.group(2) else ''
            self.analysis_results['macros_defined'][macro_name] = {
                'line': line_num,
                'parameters': params
            }
        
        # Macro calls
        macro_calls = re.findall(r'%([A-Z_][A-Z0-9_]*)\b', line)
        system_macros = {'MACRO', 'MEND', 'LET', 'IF', 'DO', 'END', 'EVAL', 'STR', 'QUOTE', 'SCAN', 'SUBSTR', 'INCLUDE'}
        for macro_name in macro_calls:
            if macro_name not in system_macros:
                self.analysis_results['macros_called'][macro_name].append(line_num)
    
    def _analyze_sql_operations(self, line, line_num):
        """Analyze SQL operations within PROC SQL"""
        sql_operations = [
            'SELECT', 'CREATE', 'INSERT', 'UPDATE', 'DELETE', 'ALTER', 'DROP',
            'FROM', 'WHERE', 'GROUP BY', 'HAVING', 'ORDER BY', 'UNION', 'JOIN'
        ]
        
        for operation in sql_operations:
            spaced_pattern = operation.replace(' ', r'\s+')
            word_boundary_pattern = rf'\b{spaced_pattern}\b'
            
            if re.search(word_boundary_pattern, line):
                operation_key = operation.replace(' ', '_')
                self.analysis_results['sql_statements'][operation_key].append(line_num)
    
    def _analyze_control_structures(self, line, line_num):
        """Analyze control structures"""
        control_patterns = [
            ('IF_THEN', r'\bIF\s+.*\bTHEN\b'),
            ('DO_LOOP', r'\bDO\b.*(?:TO|WHILE|UNTIL)'),
            ('ARRAY', r'\bARRAY\s+[A-Z_][A-Z0-9_]*'),
            ('FORMAT', r'\bFORMAT\s+'),
            ('LENGTH', r'\bLENGTH\s+'),
            ('LABEL', r'\bLABEL\s+'),
            ('RETAIN', r'\bRETAIN\s+'),
            ('OUTPUT', r'\bOUTPUT\s*;'),
            ('RETURN', r'\bRETURN\s*;'),
            ('DELETE', r'\bDELETE\s*;')
        ]
        
        for structure, pattern in control_patterns:
            if re.search(pattern, line):
                self.analysis_results['control_structures'][structure].append(line_num)
    
    def _analyze_file_operations(self, line, line_num):
        """Analyze file operations"""
        # LIBNAME statements
        lib_match = re.search(r'LIBNAME\s+([A-Z_][A-Z0-9_]*)', line)
        if lib_match:
            lib_name = lib_match.group(1)
            self.analysis_results['libraries_defined'][lib_name].append(line_num)
        
        # File operations
        file_operations = [
            ('INFILE', r'\bINFILE\s+'),
            ('FILE', r'\bFILE\s+'),
            ('FILENAME', r'\bFILENAME\s+'),
            ('PUT', r'\bPUT\s+'),
            ('INPUT', r'\bINPUT\s+')
        ]
        
        for operation, pattern in file_operations:
            if re.search(pattern, line):
                self.analysis_results['file_operations'][operation].append(line_num)
    
    def _analyze_variables(self, line, line_num):
        """Analyze variable operations"""
        var_operations = [
            ('KEEP', r'KEEP\s+([A-Z_][A-Z0-9_]*(?:\s+[A-Z_][A-Z0-9_]*)*?)(?:\s|;|$)'),
            ('DROP', r'DROP\s+([A-Z_][A-Z0-9_]*(?:\s+[A-Z_][A-Z0-9_]*)*?)(?:\s|;|$)'),
            ('VAR', r'VAR\s+([A-Z_][A-Z0-9_]*(?:\s+[A-Z_][A-Z0-9_]*)*?)(?:\s|;|$)'),
            ('BY', r'BY\s+([A-Z_][A-Z0-9_]*(?:\s+[A-Z_][A-Z0-9_]*)*?)(?:\s|;|$)'),
            ('CLASS', r'CLASS\s+([A-Z_][A-Z0-9_]*(?:\s+[A-Z_][A-Z0-9_]*)*?)(?:\s|;|$)')
        ]
        
        for operation, pattern in var_operations:
            matches = re.findall(pattern, line)
            for match in matches:
                variables = match.split()
                for var in variables:
                    self.analysis_results['variables_used'][f"{operation}_{var}"].append(line_num)
    
    def _analyze_timeframes(self, line, line_num):
        """Extract timestamps and timeframes"""
        timestamp_patterns = [
            r'\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}',
            r'\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}',
            r'\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2}',
            r'\d{2}-\d{2}-\d{4}\s+\d{2}:\d{2}:\d{2}'
        ]
        
        for pattern in timestamp_patterns:
            timestamps = re.findall(pattern, line)
            for timestamp in timestamps:
                if not self.analysis_results['timeframe_start']:
                    self.analysis_results['timeframe_start'] = {'timestamp': timestamp, 'line': line_num}
                elif timestamp < self.analysis_results['timeframe_start']['timestamp']:
                    self.analysis_results['timeframe_start'] = {'timestamp': timestamp, 'line': line_num}
                
                if not self.analysis_results['timeframe_end']:
                    self.analysis_results['timeframe_end'] = {'timestamp': timestamp, 'line': line_num}
                elif timestamp > self.analysis_results['timeframe_end']['timestamp']:
                    self.analysis_results['timeframe_end'] = {'timestamp': timestamp, 'line': line_num}
    
    def _calculate_complexity_metrics(self):
        """Calculate code complexity metrics"""
        total_functions = len(self.analysis_results['function_blocks'])
        total_datasets = len(self.analysis_results['datasets_created']) + len(self.analysis_results['datasets_used'])
        total_procedures = len(self.analysis_results['procedures_used'])
        total_macros = len(self.analysis_results['macros_defined']) + len(self.analysis_results['macros_called'])
        total_includes = len(self.analysis_results['include_files'])
        total_sys_functions = len(self.analysis_results['system_functions'])
        
        self.analysis_results['code_complexity'] = {
            'total_functions': total_functions,
            'total_datasets': total_datasets,
            'total_procedures': total_procedures,
            'total_macros': total_macros,
            'total_includes': total_includes,
            'total_sys_functions': total_sys_functions,
            'complexity_score': total_functions + total_procedures + total_macros + total_includes,
            'total_lines': len(self.analysis_results['line_analysis'])
        }

class SASReportGenerator:
    def __init__(self, analysis_results, source_file):
        self.results = analysis_results
        self.source_file = source_file
        self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    def generate_summary_report(self):
        """Generate enhanced summary report"""
        if not self.results:
            return "❌ No analysis results available"
        
        complexity = self.results['code_complexity']
        
        report = f"""
🔍 **SAS CODE ANALYSIS SUMMARY REPORT**
{'='*70}
📁 **SOURCE FILE:** {self.source_file}
📅 **ANALYSIS DATE:** {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

📊 **CODE METRICS:**
   • Total Lines: {complexity['total_lines']}
   • Complexity Score: {complexity['complexity_score']}
   • Function Blocks: {complexity['total_functions']}
   • Procedures: {complexity['total_procedures']}
   • Macros: {complexity['total_macros']}
   • Include Files: {complexity['total_includes']}
   • System Functions: {complexity['total_sys_functions']}

📝 **KEY COMPONENTS:**
   • Datasets Created: {len(self.results['datasets_created'])}
   • Datasets Used: {len(self.results['datasets_used'])}
   • Libraries Defined: {len(self.results['libraries_defined'])}
   • SQL Statements: {len(self.results['sql_statements'])}
   • Control Structures: {len(self.results['control_structures'])}
   • CALL Routines: {len(self.results['call_routines'])}
   • Hash Objects: {len(self.results['hash_objects'])}
   • ODS Statements: {len(self.results['ods_statements'])}
"""
        
        # Add timeframe if available
        if self.results['timeframe_start'] and self.results['timeframe_end']:
            report += f"""
⏰ **EXECUTION TIMEFRAME:**
   • Start: {self.results['timeframe_start']['timestamp']} (Line {self.results['timeframe_start']['line']})
   • End: {self.results['timeframe_end']['timestamp']} (Line {self.results['timeframe_end']['line']})
"""
        
        return report
    
    def generate_detailed_report(self):
        """Generate comprehensive detailed report"""
        if not self.results:
            return "❌ No analysis results available"
        
        report = f"""
🔍 **COMPREHENSIVE SAS CODE ANALYSIS REPORT**
{'='*80}
📁 **SOURCE FILE:** {self.source_file}
📅 **ANALYSIS DATE:** {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

{self.generate_summary_report()}

🔧 **FUNCTION BLOCKS (Start → End Lines):**
"""
        
        # Function blocks with start and end lines
        for block in self.results['function_blocks']:
            duration = block['end_line'] - block['start_line'] + 1
            report += f"   • {block['name']}: Lines {block['start_line']} → {block['end_line']} ({duration} lines)\n"
        
        # Include files
        if self.results['include_files']:
            report += f"\n📂 **INCLUDE FILES:**\n"
            for include_file, lines in self.results['include_files'].items():
                report += f"   • {include_file}: Lines {', '.join(map(str, lines))}\n"
        
        # System functions
        if self.results['system_functions']:
            report += f"\n🔨 **SYSTEM FUNCTIONS:**\n"
            for func, lines in list(self.results['system_functions'].items())[:15]:
                report += f"   • {func}(): Lines {', '.join(map(str, lines))}\n"
            if len(self.results['system_functions']) > 15:
                report += f"   ... and {len(self.results['system_functions']) - 15} more system functions\n"
        
        # CALL routines
        if self.results['call_routines']:
            report += f"\n📞 **CALL ROUTINES:**\n"
            for routine, lines in self.results['call_routines'].items():
                report += f"   • {routine}: Lines {', '.join(map(str, lines))}\n"
        
        # Hash objects
        if self.results['hash_objects']:
            report += f"\n🗃️ **HASH OBJECTS:**\n"
            for hash_obj, lines in self.results['hash_objects'].items():
                report += f"   • {hash_obj}: Lines {', '.join(map(str, lines))}\n"
        
        # ODS statements
        if self.results['ods_statements']:
            report += f"\n📊 **ODS STATEMENTS:**\n"
            for ods_stmt, lines in self.results['ods_statements'].items():
                report += f"   • {ods_stmt}: Lines {', '.join(map(str, lines))}\n"
        
        # Formats
        if self.results['formats']:
            report += f"\n🎨 **FORMATS:**\n"
            for fmt, lines in self.results['formats'].items():
                report += f"   • {fmt}: Lines {', '.join(map(str, lines))}\n"
        
        # Datasets
        if self.results['datasets_created']:
            report += f"\n📝 **DATASETS CREATED:**\n"
            for dataset, lines in self.results['datasets_created'].items():
                report += f"   • {dataset}: Lines {', '.join(map(str, lines))}\n"
        
        if self.results['datasets_used']:
            report += f"\n📖 **DATASETS USED:**\n"
            for usage, lines in self.results['datasets_used'].items():
                report += f"   • {usage}: Lines {', '.join(map(str, lines))}\n"
        
        # Procedures
        if self.results['procedures_used']:
            report += f"\n⚙️ **PROCEDURES:**\n"
            for proc, instances in self.results['procedures_used'].items():
                lines_info = [f"{inst['line']} (data={inst['dataset']})" for inst in instances]
                report += f"   • PROC {proc}: Lines {', '.join(lines_info)}\n"
        
        # Macros
        if self.results['macros_defined']:
            report += f"\n🔨 **MACRO DEFINITIONS:**\n"
            for macro, info in self.results['macros_defined'].items():
                report += f"   • %{macro}{info['parameters']}: Line {info['line']}\n"
        
        if self.results['macros_called']:
            report += f"\n📞 **MACRO CALLS:**\n"
            for macro, lines in self.results['macros_called'].items():
                report += f"   • %{macro}: Lines {', '.join(map(str, lines))}\n"
        
        # SQL statements
        if self.results['sql_statements']:
            report += f"\n🗃️ **SQL OPERATIONS:**\n"
            for operation, lines in self.results['sql_statements'].items():
                report += f"   • {operation}: Lines {', '.join(map(str, lines))}\n"
        
        # Control structures
        if self.results['control_structures']:
            report += f"\n🔄 **CONTROL STRUCTURES:**\n"
            for structure, lines in self.results['control_structures'].items():
                report += f"   • {structure}: Lines {', '.join(map(str, lines))}\n"
        
        # Libraries
        if self.results['libraries_defined']:
            report += f"\n📁 **LIBRARIES:**\n"
            for lib, lines in self.results['libraries_defined'].items():
                report += f"   • {lib}: Lines {', '.join(map(str, lines))}\n"
        
        # Variables (showing top 10)
        if self.results['variables_used']:
            report += f"\n🔤 **VARIABLE OPERATIONS (Top 10):**\n"
            var_items = list(self.results['variables_used'].items())[:10]
            for var_op, lines in var_items:
                report += f"   • {var_op}: Lines {', '.join(map(str, lines))}\n"
            if len(self.results['variables_used']) > 10:
                report += f"   ... and {len(self.results['variables_used']) - 10} more variable operations\n"
        
        return report
    
    def save_reports(self, output_dir='reports'):
        """Save both summary and detailed reports to files"""
        # Create reports directory if it doesn't exist
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"✅ Created reports directory: {output_dir}")
        
        # Generate reports
        summary_report = self.generate_summary_report()
        detailed_report = self.generate_detailed_report()
        
        # Get base filename without extension
        base_filename = os.path.splitext(os.path.basename(self.source_file))[0]
        
        # Save summary report
        summary_filename = f"{output_dir}/{base_filename}_summary_{self.timestamp}.txt"
        with open(summary_filename, 'w', encoding='utf-8') as f:
            f.write(summary_report)
        print(f"✅ Summary report saved: {summary_filename}")
        
        # Save detailed report
        detailed_filename = f"{output_dir}/{base_filename}_detailed_{self.timestamp}.txt"
        with open(detailed_filename, 'w', encoding='utf-8') as f:
            f.write(detailed_report)
        print(f"✅ Detailed report saved: {detailed_filename}")
        
        return summary_filename, detailed_filename

def analyze_sas_file(input_file_path, output_dir='reports', show_console_output=True):
    """
    Main function to analyze SAS file and generate reports
    
    Args:
        input_file_path: Path to input SAS .txt file
        output_dir: Directory to save report files
        show_console_output: Whether to print reports to console
    
    Returns:
        Tuple of (analysis_results, summary_report_file, detailed_report_file)
    """
    
    print("🚀 Starting SAS Code Analysis...")
    print("="*60)
    
    # Step 1: Analyze the file
    analyzer = EnhancedSASAnalyzer()
    results = analyzer.analyze_file(input_file_path)
    
    if not results:
        print("❌ Analysis failed")
        return None, None, None
    
    # Step 2: Generate reports
    report_generator = SASReportGenerator(results, input_file_path)
    
    # Step 3: Save reports to files
    summary_file, detailed_file = report_generator.save_reports(output_dir)
    
    # Step 4: Show console output if requested
    if show_console_output:
        print("\n" + "="*60)
        print("📊 CONSOLE OUTPUT - SUMMARY REPORT")
        print("="*60)
        print(report_generator.generate_summary_report())
        
        print("\n" + "="*60)
        print("📋 CONSOLE OUTPUT - DETAILED REPORT")
        print("="*60)
        print(report_generator.generate_detailed_report())
    
    print(f"\n✅ Analysis completed successfully!")
    print(f"📁 Reports saved in: {output_dir}/")
    
    return results, summary_file, detailed_file

# Main execution
if __name__ == "__main__":
    print("🔍 SAS Code Analyzer")
    print("="*50)
    
    # Get input file from user
    input_file = input("📁 Enter path to your SAS .txt file: ").strip()
    
    if not input_file:
        print("❌ No file specified. Exiting.")
        exit()
    
    if not os.path.exists(input_file):
        print(f"❌ File not found: {input_file}")
        exit()
    
    # Optional: Get output directory
    output_directory = input("📂 Enter reports output directory (press Enter for 'reports'): ").strip()
    if not output_directory:
        output_directory = 'reports'
    
    # Analyze the file
    try:
        results, summary_file, detailed_file = analyze_sas_file(
            input_file_path=input_file,
            output_dir=output_directory,
            show_console_output=True
        )
        
        print("\n🎉 Analysis Complete!")
        print(f"📈 Summary Report: {summary_file}")
        print(f"📊 Detailed Report: {detailed_file}")
        
    except Exception as e:
        print(f"❌ Error during analysis: {e}")

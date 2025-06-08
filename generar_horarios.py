import pandas as pd
import os
from typing import Dict, List, Tuple

class LaboratoryScheduleGenerator:
    def __init__(self):
        """
        Initialize the Laboratory Schedule Generator with default configurations.
        """
        # Lab name mapping - EDIT THIS DICTIONARY to match your lab names
        # Key: Lab name in reporte_ocupacion file
        # Value: Lab name for output headers
        self.lab_mapping = {
            "LABORATORIO GEIO CAP(25)": "GEIO (321) TECHNE",
            "SALA DE SOFTWARE DE TECNOLOGIA E INGENIERIA DE PRODUCCION A CAP(17)": "Sala de Software A - 16 EST - 416- TECHNE",
            "SALA DE SOFTWARE DE TECNOLOGIA E NGENIERIA DE PRODUCCION B CAP(25)": "Sala de Software B - 24 EST - 417 TECHNE", # Corrected typo "NGENIERIA" if it exists
            "LABORATORIO HAS CAP(22)": "HAS-200 (317) TECHNE",
            "LABORATORIO FMS CAP(18)": "FMS-200 (320) TECHNE",
            "LABORATORIO DE PROCESOS DE TRANSFORMACIÓN MECÁNICA": "LABORATORIO DE PROCESOS DE TRANSFORMACIÓN BLOQUE 1-102",
            # Add more mappings as needed
            # 'SOURCE_LAB_NAME': 'OUTPUT_LAB_NAME',
        }
        
        # Days of the week in order
        self.days = ['LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADOS']
        
        # Time slots in order (6AM to 5PM)
        self.time_slots = [
            '6AM-7AM', '7AM-8AM', '8AM-9AM', '9AM-10AM', '10AM-11AM',
            '11AM-12M', '12M-1PM', '1PM-2PM', '2PM-3PM', '3PM-4PM',
            '4PM-5PM', '5PM-6PM', '6PM-7PM', '7PM-8PM', '8PM-9PM', '9PM-10PM'
        ]
        
        # Expected columns in the input file
        self.input_columns = [
            'Periodo', 'Día', 'Hora', 'Asignatura', 'Grupo', 
            'Proyecto', 'Salón', 'Área', 'Edificio', 'Sede', 
            'Inscritos', 'Docente'
        ]

    def update_lab_mapping(self, new_mapping: Dict[str, str]):
        """
        Update the lab name mapping dictionary.
        
        Args:
            new_mapping: Dictionary with lab name mappings
        """
        self.lab_mapping.update(new_mapping)
        print("Lab mapping updated successfully!")

    def read_occupation_report(self, file_path: str) -> pd.DataFrame:
        """
        Read the occupation report Excel file.
        
        Args:
            file_path: Path to the reporte_ocupacion.xlsx file
            
        Returns:
            DataFrame with the occupation data
        """
        try:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File not found: {file_path}")
                
            df = pd.read_excel(file_path)
            
            # Validate columns
            if len(df.columns) != len(self.input_columns):
                print(f"Warning: Expected {len(self.input_columns)} columns, got {len(df.columns)}")
                print(f"Expected: {self.input_columns}")
                print(f"Found: {list(df.columns)}")
            
            # Rename columns to match expected names
            df.columns = self.input_columns[:len(df.columns)]
            
            print(f"Successfully loaded {len(df)} records from {file_path}")
            return df
            
        except Exception as e:
            raise Exception(f"Error reading occupation report: {str(e)}")

    def filter_mapped_labs(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Filter the dataframe to only include labs that are in our mapping.
        
        Args:
            df: Input dataframe
            
        Returns:
            Filtered dataframe with only mapped labs
        """
        mapped_labs = list(self.lab_mapping.keys())
        filtered_df = df[df['Salón'].isin(mapped_labs)].copy()
        
        print(f"Filtered to {len(filtered_df)} records for mapped laboratories")
        print(f"Mapped labs found: {filtered_df['Salón'].unique().tolist()}")
        
        return filtered_df

    def group_consecutive_hours(self, df: pd.DataFrame) -> List[Dict]:
        """
        Group consecutive hour entries into single class sessions.
        Also handles single-hour classes and non-consecutive classes.
        
        Args:
            df: Filtered dataframe
            
        Returns:
            List of class session dictionaries
        """
        classes = []
        
        # Group by all fields except hour
        group_fields = ['Día', 'Asignatura', 'Grupo', 'Proyecto', 'Salón', 'Docente']
        grouped = df.groupby(group_fields)
        
        for group_key, group_df in grouped:
            # Sort by hour to ensure proper consecutive ordering
            group_df = group_df.sort_values('Hora')
            hours = group_df['Hora'].tolist()
            processed_hours = set()
            
            # Find consecutive pairs first
            i = 0
            while i < len(hours) - 1:
                if hours[i] in processed_hours:
                    i += 1
                    continue
                    
                current_hour = hours[i]
                next_hour = hours[i + 1]
                
                # Check if they are consecutive
                if self.are_consecutive_hours(current_hour, next_hour):
                    class_info = {
                        'day': group_key[0],
                        'start_hour': current_hour,
                        'end_hour': next_hour,
                        'subject': group_key[1],
                        'group': group_key[2],
                        'project': group_key[3],
                        'lab': group_key[4],
                        'teacher': group_key[5],
                        'is_two_hour': True
                    }
                    classes.append(class_info)
                    processed_hours.add(current_hour)
                    processed_hours.add(next_hour)
                    i += 2  # Skip the next hour as it's already processed
                else:
                    i += 1
            
            # Handle remaining single hours or non-consecutive classes
            for hour in hours:
                if hour not in processed_hours:
                    class_info = {
                        'day': group_key[0],
                        'start_hour': hour,
                        'end_hour': None,
                        'subject': group_key[1],
                        'group': group_key[2],
                        'project': group_key[3],
                        'lab': group_key[4],
                        'teacher': group_key[5],
                        'is_two_hour': False
                    }
                    classes.append(class_info)
        
        print(f"Found {len(classes)} class sessions (including single-hour classes)")
        return classes

    def are_consecutive_hours(self, hour1: str, hour2: str) -> bool:
        """
        Check if two hours are consecutive in our time slot sequence.
        
        Args:
            hour1: First hour
            hour2: Second hour
            
        Returns:
            True if consecutive, False otherwise
        """
        try:
            idx1 = self.time_slots.index(hour1)
            idx2 = self.time_slots.index(hour2)
            return idx2 == idx1 + 1
        except ValueError:
            return False

    def create_schedule_matrix(self, classes: List[Dict]) -> pd.DataFrame:
        """
        Create the output schedule matrix with the exact template structure.
        
        Args:
            classes: List of class session dictionaries
            
        Returns:
            DataFrame with the schedule matrix
        """
        # Get unique labs from our mapping (output names)
        output_labs = list(set(self.lab_mapping.values()))
        output_labs.sort()  # Sort for consistent ordering
        
        # Create columns: Dia, Hora, then triplets for each lab (Subject, Group, Teacher/Project)
        columns = ['Dia', 'Hora']
        for lab in output_labs:
            columns.extend([f"{lab}_subject", f"{lab}_group", f"{lab}_teacher_project"])
        
        # Create rows for each day and time slot, plus separation rows
        rows = []
        for day_idx, day in enumerate(self.days):
            for time_slot in self.time_slots:
                row = {'Dia': day, 'Hora': time_slot}
                # Initialize all lab columns as empty
                for lab in output_labs:
                    row[f"{lab}_subject"] = ''
                    row[f"{lab}_group"] = ''
                    row[f"{lab}_teacher_project"] = ''
                rows.append(row)
            
            # Add separation row after each day (except the last one)
            if day_idx < len(self.days) - 1:
                separator_row = {'Dia': '', 'Hora': ''}
                for lab in output_labs:
                    separator_row[f"{lab}_subject"] = ''
                    separator_row[f"{lab}_group"] = ''
                    separator_row[f"{lab}_teacher_project"] = ''
                rows.append(separator_row)
        
        # Create the base dataframe
        schedule_df = pd.DataFrame(rows, columns=columns)
        
        # Fill in the class information
        for class_info in classes:
            day = class_info['day']
            start_hour = class_info['start_hour']
            
            # Map lab name to output name
            output_lab = self.lab_mapping.get(class_info['lab'])
            if not output_lab:
                continue
            
            if class_info['is_two_hour']:
                # Two-hour class: traditional format
                end_hour = class_info['end_hour']
                
                # First hour: Subject in first column, Group in second column
                start_mask = (schedule_df['Dia'] == day) & (schedule_df['Hora'] == start_hour)
                if start_mask.any():
                    idx = schedule_df[start_mask].index[0]
                    schedule_df.loc[idx, f"{output_lab}_subject"] = class_info['subject']
                    schedule_df.loc[idx, f"{output_lab}_group"] = class_info['group']
                
                # Second hour: Teacher in first column, Project in second column
                end_mask = (schedule_df['Dia'] == day) & (schedule_df['Hora'] == end_hour)
                if end_mask.any():
                    idx = schedule_df[end_mask].index[0]
                    schedule_df.loc[idx, f"{output_lab}_subject"] = class_info['teacher']
                    schedule_df.loc[idx, f"{output_lab}_group"] = class_info['project']
            else:
                # Single-hour class: put all info in one row
                start_mask = (schedule_df['Dia'] == day) & (schedule_df['Hora'] == start_hour)
                if start_mask.any():
                    idx = schedule_df[start_mask].index[0]
                    schedule_df.loc[idx, f"{output_lab}_subject"] = class_info['subject']
                    schedule_df.loc[idx, f"{output_lab}_group"] = class_info['group']
                    schedule_df.loc[idx, f"{output_lab}_teacher_project"] = f"{class_info['teacher']} | {class_info['project']}"
        
        return schedule_df

    def format_output_headers(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Format the output headers to match the template structure.
        
        Args:
            df: Schedule dataframe
            
        Returns:
            DataFrame with properly formatted headers
        """
        # Create a copy to avoid modifying the original
        output_df = df.copy()
        
        # Create new column names for better readability
        new_columns = []
        lab_names = []
        
        for col in output_df.columns:
            if col in ['Dia', 'Hora']:
                new_columns.append(col)
            elif col.endswith('_subject'):
                lab_name = col.replace('_subject', '')
                new_columns.append(f"{lab_name} - Asignatura")
                if lab_name not in lab_names:
                    lab_names.append(lab_name)
            elif col.endswith('_group'):
                lab_name = col.replace('_group', '')
                new_columns.append(f"{lab_name} - Grupo")
            elif col.endswith('_teacher_project'):
                lab_name = col.replace('_teacher_project', '')
                new_columns.append(f"{lab_name} - Docente/Proyecto")
        
        output_df.columns = new_columns
        
        return output_df, lab_names

    def save_schedule(self, df: pd.DataFrame, output_path: str):
        """
        Save the schedule to an Excel file with proper formatting and day separations.
        
        Args:
            df: Schedule dataframe
            output_path: Output file path
        """
        try:
            from openpyxl import Workbook
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
            from openpyxl.utils.dataframe import dataframe_to_rows
            
            # Create workbook and worksheet
            wb = Workbook()
            ws = wb.active
            ws.title = "Horario Laboratorios"
            
            # Write the dataframe to the worksheet
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)
            
            # Format headers
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # Format day separation rows (empty rows between days)
            separator_fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
            
            for row in range(2, ws.max_row + 1):
                # Check if this is a separator row (empty Dia and Hora)
                if not ws.cell(row=row, column=1).value and not ws.cell(row=row, column=2).value:
                    for col in range(1, ws.max_column + 1):
                        ws.cell(row=row, column=col).fill = separator_fill
            
            # Add borders to all cells
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            for row in range(1, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).border = thin_border
                    ws.cell(row=row, column=col).alignment = Alignment(
                        horizontal="center", 
                        vertical="center",
                        wrap_text=True
                    )
            
            # Auto-adjust column widths
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
                ws.column_dimensions[column_letter].width = adjusted_width
            
            # Save the workbook
            wb.save(output_path)
            
            print(f"Schedule saved successfully to: {output_path}")
            
        except Exception as e:
            # Fallback to basic pandas save if openpyxl formatting fails
            print(f"Advanced formatting failed, saving basic version: {str(e)}")
            df.to_excel(output_path, index=False, sheet_name='Horario')
            print(f"Basic schedule saved to: {output_path}")

    def generate_schedule(self, input_file: str, output_file: str):
        """
        Main method to generate the complete schedule.
        
        Args:
            input_file: Path to reporte_ocupacion.xlsx
            output_file: Path for output HORARIO_LABORATORIOS.xlsx
        """
        print("Starting laboratory schedule generation...")
        print(f"Input file: {input_file}")
        print(f"Output file: {output_file}")
        print(f"Lab mappings: {self.lab_mapping}")
        
        try:
            # Step 1: Read the occupation report
            df = self.read_occupation_report(input_file)
            
            # Step 2: Filter for mapped laboratories
            filtered_df = self.filter_mapped_labs(df)
            
            if filtered_df.empty:
                print("Warning: No data found for mapped laboratories!")
                return
            
            # Step 3: Group consecutive hours into classes
            classes = self.group_consecutive_hours(filtered_df)
            
            if not classes:
                print("Warning: No complete class sessions found!")
                return
            
            # Step 4: Create the schedule matrix
            schedule_df = self.create_schedule_matrix(classes)
            
            # Step 5: Format headers
            formatted_df, lab_names = self.format_output_headers(schedule_df)
            
            # Step 6: Save the output
            self.save_schedule(formatted_df, output_file)
            
            print("Laboratory schedule generation completed successfully!")
            
        except Exception as e:
            print(f"Error during schedule generation: {str(e)}")
            raise

def main():
    """
    Example usage of the Laboratory Schedule Generator
    """
    # Create the generator instance
    generator = LaboratoryScheduleGenerator()
    
    # Optional: Update lab mappings if needed
    additional_mappings = {
        # 'NEW_SOURCE_LAB': 'NEW_OUTPUT_LAB',
    }
    if additional_mappings:
        generator.update_lab_mapping(additional_mappings)
    
    # File paths
    input_file = "reporte_ocupacion.xlsx"
    output_file = "HORARIO_LABORATORIOS.xlsx"
    
    # Generate the schedule
    try:
        generator.generate_schedule(input_file, output_file)
    except Exception as e:
        print(f"Failed to generate schedule: {str(e)}")

if __name__ == "__main__":
    main()
import pandas as pd
import numpy as np
from openpyxl import load_workbook

def calculate_detailed_august_metrics():
    """Calculate detailed August metrics for the specific requirements"""
    
    try:
        # Read the Excel file
        file_path = "August Export_SD 2 Sept_modified.xlsx"
        print(f"Reading August data from: {file_path}")
        
        # Read August sheet
        august_df = pd.read_excel(file_path, sheet_name='August')
        
        # Clean column names
        august_df.columns = august_df.columns.str.strip()
        
        print("\n" + "="*80)
        print("📊 DETAILED AUGUST METRICS CALCULATION")
        print("="*80)
        
        # 1. Average VWE modules commenced per student
        print("\n🔍 1. AVERAGE VWE MODULES COMMENCED PER STUDENT")
        print("-" * 60)
        
        vwe_column = 'Virtual Work Experience'
        if vwe_column in august_df.columns:
            vwe_data = august_df[vwe_column].dropna()
            total_students = len(august_df)
            students_with_vwe = len(vwe_data)
            
            if pd.api.types.is_numeric_dtype(vwe_data):
                avg_vwe = vwe_data.mean()
                print(f"Total students in August: {total_students}")
                print(f"Students with VWE data: {students_with_vwe}")
                print(f"Average VWE modules commenced per student: {avg_vwe:.2f}")
                print(f"VWE data range: {vwe_data.min()} to {vwe_data.max()}")
                print(f"VWE data distribution: {vwe_data.value_counts().sort_index().to_dict()}")
            else:
                print("VWE data is not numeric")
        else:
            print(f"VWE column '{vwe_column}' not found")
        
        # 2. Average industry-based modules completed per student
        print("\n🏭 2. AVERAGE INDUSTRY-BASED MODULES COMPLETED PER STUDENT")
        print("-" * 60)
        
        industry_column = 'Industries'
        if industry_column in august_df.columns:
            industry_data = august_df[industry_column].dropna()
            students_with_industry = len(industry_data)
            
            print(f"Total students in August: {total_students}")
            print(f"Students with industry data: {students_with_industry}")
            
            # Analyze industry data structure
            sample_industry = industry_data.iloc[0] if len(industry_data) > 0 else None
            print(f"Sample industry data: {sample_industry}")
            
            if sample_industry and isinstance(sample_industry, str):
                # Count modules per student
                module_counts = []
                for value in industry_data:
                    if pd.notna(value) and isinstance(value, str):
                        # Split by pipe and count non-empty modules
                        modules = [m.strip() for m in str(value).split('|') if m.strip() and m.strip() != "'"]
                        module_counts.append(len(modules))
                    else:
                        module_counts.append(0)
                
                if module_counts:
                    avg_industry_modules = np.mean(module_counts)
                    print(f"Average industry-based modules completed per student: {avg_industry_modules:.2f}")
                    print(f"Module count distribution: {np.unique(module_counts, return_counts=True)}")
                    print(f"Module count range: {min(module_counts)} to {max(module_counts)}")
                else:
                    print("No valid module counts found")
            else:
                print("Industry data format not recognized")
        else:
            print(f"Industry column '{industry_column}' not found")
        
        # 3. Average number of modules engaged with per session per student
        print("\n📚 3. AVERAGE MODULES ENGAGED WITH PER SESSION PER STUDENT")
        print("-" * 60)
        
        web_sessions_col = 'Web sessions'
        person_tag_col = 'Person tag'
        
        if web_sessions_col in august_df.columns and person_tag_col in august_df.columns:
            web_sessions = august_df[web_sessions_col].dropna()
            person_tags = august_df[person_tag_col].dropna()
            
            print(f"Total students in August: {total_students}")
            print(f"Students with web sessions data: {len(web_sessions)}")
            print(f"Students with person tag data: {len(person_tags)}")
            
            # Calculate modules engaged per session
            modules_per_session = []
            valid_students = 0
            
            for i in range(len(august_df)):
                web_session_count = august_df.iloc[i][web_sessions_col]
                person_tag_value = august_df.iloc[i][person_tag_col]
                
                if pd.notna(web_session_count) and pd.notna(person_tag_value) and web_session_count > 0:
                    if isinstance(person_tag_value, str):
                        # Count different engagement types in person tag
                        engagements = [e.strip() for e in str(person_tag_value).split(',') if e.strip()]
                        modules_per_session.append(len(engagements) / web_session_count)
                        valid_students += 1
            
            if modules_per_session:
                avg_modules_per_session = np.mean(modules_per_session)
                print(f"Students with valid session and engagement data: {valid_students}")
                print(f"Average modules engaged with per session per student: {avg_modules_per_session:.2f}")
                print(f"Modules per session range: {min(modules_per_session):.2f} to {max(modules_per_session):.2f}")
                print(f"Modules per session distribution: {np.percentile(modules_per_session, [25, 50, 75])}")
            else:
                print("No valid modules per session data found")
        else:
            print("Required columns not found")
        
        # Final Summary
        print("\n" + "="*80)
        print("📋 FINAL AUGUST METRICS SUMMARY")
        print("="*80)
        
        # Calculate final metrics
        final_metrics = {}
        
        # VWE metric
        if vwe_column in august_df.columns:
            vwe_data = august_df[vwe_column].dropna()
            if pd.api.types.is_numeric_dtype(vwe_data):
                final_metrics['VWE'] = vwe_data.mean()
        
        # Industry metric
        if industry_column in august_df.columns:
            industry_data = august_df[industry_column].dropna()
            if len(industry_data) > 0:
                module_counts = []
                for value in industry_data:
                    if pd.notna(value) and isinstance(value, str):
                        modules = [m.strip() for m in str(value).split('|') if m.strip() and m.strip() != "'"]
                        module_counts.append(len(modules))
                    else:
                        module_counts.append(0)
                if module_counts:
                    final_metrics['Industry'] = np.mean(module_counts)
        
        # Modules per session metric
        if web_sessions_col in august_df.columns and person_tag_col in august_df.columns:
            modules_per_session = []
            for i in range(len(august_df)):
                web_session_count = august_df.iloc[i][web_sessions_col]
                person_tag_value = august_df.iloc[i][person_tag_col]
                
                if pd.notna(web_session_count) and pd.notna(person_tag_value) and web_session_count > 0:
                    if isinstance(person_tag_value, str):
                        engagements = [e.strip() for e in str(person_tag_value).split(',') if e.strip()]
                        modules_per_session.append(len(engagements) / web_session_count)
            
            if modules_per_session:
                final_metrics['Modules_Per_Session'] = np.mean(modules_per_session)
        
        # Display final formatted results
        print("\n🎯 FINAL RESULTS FOR AUGUST:")
        print("-" * 40)
        print(f"Total students in August: {total_students}")
        print(f"Average VWE modules commenced per student: {final_metrics.get('VWE', 'N/A'):.2f}" if 'VWE' in final_metrics else "Average VWE modules commenced per student: N/A")
        print(f"Average industry-based modules completed per student: {final_metrics.get('Industry', 'N/A'):.2f}" if 'Industry' in final_metrics else "Average industry-based modules completed per student: N/A")
        print(f"Average modules engaged with per session per student: {final_metrics.get('Modules_Per_Session', 'N/A'):.2f}" if 'Modules_Per_Session' in final_metrics else "Average modules engaged with per session per student: N/A")
        
        return final_metrics
        
    except Exception as e:
        print(f"Error calculating detailed metrics: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    final_metrics = calculate_detailed_august_metrics()
    if final_metrics:
        print(f"\nDetailed analysis completed successfully!")
        print(f"Final metrics calculated: {final_metrics}")

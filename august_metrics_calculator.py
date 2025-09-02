import pandas as pd
import numpy as np
from openpyxl import load_workbook

def calculate_august_metrics():
    """Calculate August metrics for VWE modules, industry modules, and engagement per session"""
    
    try:
        # Read the Excel file
        file_path = "August Export_SD 2 Sept_modified.xlsx"
        print(f"Reading August data from: {file_path}")
        
        # Read August sheet
        august_df = pd.read_excel(file_path, sheet_name='August')
        print(f"August data shape: {august_df.shape}")
        
        # Clean column names
        august_df.columns = august_df.columns.str.strip()
        print(f"August columns: {list(august_df.columns)}")
        
        # Display sample data to understand structure
        print("\nSample August data:")
        print(august_df.head())
        
        print("\n" + "="*80)
        print("📊 AUGUST METRICS CALCULATION")
        print("="*80)
        
        # 1. Calculate Average VWE modules commenced per student
        print("\n🔍 1. AVERAGE VWE MODULES COMMENCED PER STUDENT")
        print("-" * 60)
        
        vwe_column = 'Virtual Work Experience'
        if vwe_column in august_df.columns:
            vwe_data = august_df[vwe_column].dropna()
            print(f"VWE column found: {vwe_column}")
            print(f"Students with VWE data: {len(vwe_data)}")
            print(f"VWE data sample: {vwe_data.head().tolist()}")
            
            # Check if VWE is numeric or categorical
            if pd.api.types.is_numeric_dtype(vwe_data):
                avg_vwe = vwe_data.mean()
                print(f"Average VWE modules commenced per student: {avg_vwe:.2f}")
            else:
                # If categorical, count unique values
                unique_vwe = vwe_data.nunique()
                print(f"Unique VWE categories: {unique_vwe}")
                print(f"VWE categories: {vwe_data.unique()}")
                # Try to extract numeric values if possible
                try:
                    numeric_vwe = pd.to_numeric(vwe_data, errors='coerce').dropna()
                    if len(numeric_vwe) > 0:
                        avg_vwe = numeric_vwe.mean()
                        print(f"Average VWE modules commenced per student: {avg_vwe:.2f}")
                    else:
                        print("Could not convert VWE to numeric values")
                except:
                    print("Could not convert VWE to numeric values")
        else:
            print(f"VWE column '{vwe_column}' not found")
        
        # 2. Calculate Average industry-based modules completed per student
        print("\n🏭 2. AVERAGE INDUSTRY-BASED MODULES COMPLETED PER STUDENT")
        print("-" * 60)
        
        # Look for industry-related columns
        industry_columns = [col for col in august_df.columns if 'industry' in col.lower() or 'industries' in col.lower()]
        print(f"Industry-related columns found: {industry_columns}")
        
        if industry_columns:
            for col in industry_columns:
                print(f"\nAnalyzing column: {col}")
                industry_data = august_df[col].dropna()
                print(f"Students with industry data: {len(industry_data)}")
                print(f"Sample data: {industry_data.head().tolist()}")
                
                # Check if it's a list/comma-separated values
                if industry_data.dtype == 'object':
                    # Count modules per student
                    module_counts = []
                    for value in industry_data:
                        if pd.notna(value) and isinstance(value, str):
                            # Split by comma and count
                            modules = [m.strip() for m in str(value).split(',') if m.strip()]
                            module_counts.append(len(modules))
                        else:
                            module_counts.append(0)
                    
                    if module_counts:
                        avg_industry_modules = np.mean(module_counts)
                        print(f"Average industry-based modules completed per student: {avg_industry_modules:.2f}")
                        print(f"Module count distribution: {np.unique(module_counts, return_counts=True)}")
                    else:
                        print("No valid module counts found")
                else:
                    print(f"Column is numeric: {industry_data.dtype}")
                    avg_industry_modules = industry_data.mean()
                    print(f"Average industry-based modules completed per student: {avg_industry_modules:.2f}")
        else:
            print("No industry-related columns found")
        
        # 3. Calculate Average number of modules engaged with per session per student
        print("\n📚 3. AVERAGE MODULES ENGAGED WITH PER SESSION PER STUDENT")
        print("-" * 60)
        
        # Look for session and module related columns
        session_columns = [col for col in august_df.columns if 'session' in col.lower() or 'web' in col.lower()]
        print(f"Session-related columns found: {session_columns}")
        
        if session_columns:
            web_sessions_col = 'Web sessions'
            if web_sessions_col in august_df.columns:
                web_sessions = august_df[web_sessions_col].dropna()
                print(f"Web sessions data: {web_sessions.describe()}")
                
                # Look for modules engaged columns
                module_columns = [col for col in august_df.columns if 'module' in col.lower() or 'engaged' in col.lower()]
                print(f"Module engagement columns found: {module_columns}")
                
                if module_columns:
                    for col in module_columns:
                        print(f"\nAnalyzing module engagement column: {col}")
                        module_data = august_df[col].dropna()
                        print(f"Students with module data: {len(module_data)}")
                        print(f"Sample data: {module_data.head().tolist()}")
                        
                        # Try to calculate modules per session
                        if len(web_sessions) == len(module_data):
                            # Calculate modules per session for each student
                            modules_per_session = []
                            for i in range(len(web_sessions)):
                                if web_sessions.iloc[i] > 0 and pd.notna(module_data.iloc[i]):
                                    if isinstance(module_data.iloc[i], str):
                                        # Count modules in the string
                                        modules = [m.strip() for m in str(module_data.iloc[i]).split(',') if m.strip()]
                                        modules_per_session.append(len(modules) / web_sessions.iloc[i])
                                    else:
                                        modules_per_session.append(module_data.iloc[i] / web_sessions.iloc[i])
                            
                            if modules_per_session:
                                avg_modules_per_session = np.mean(modules_per_session)
                                print(f"Average modules engaged with per session per student: {avg_modules_per_session:.2f}")
                            else:
                                print("No valid modules per session data found")
                        else:
                            print("Web sessions and module data have different lengths")
                else:
                    print("No module engagement columns found")
            else:
                print(f"Web sessions column '{web_sessions_col}' not found")
        else:
            print("No session-related columns found")
        
        # Alternative approach: Look at Person tag column for engagement information
        print("\n🔍 ALTERNATIVE APPROACH: Analyzing Person tag column for engagement")
        print("-" * 60)
        
        person_tag_col = 'Person tag'
        if person_tag_col in august_df.columns:
            person_tags = august_df[person_tag_col].dropna()
            print(f"Students with Person tag data: {len(person_tags)}")
            
            # Count different types of engagement per student
            engagement_counts = []
            for tag in person_tags:
                if isinstance(tag, str):
                    # Count different engagement types
                    engagements = [e.strip() for e in str(tag).split(',') if e.strip()]
                    engagement_counts.append(len(engagements))
                else:
                    engagement_counts.append(0)
            
            if engagement_counts:
                avg_engagements = np.mean(engagement_counts)
                print(f"Average engagement types per student: {avg_engagements:.2f}")
                print(f"Engagement count distribution: {np.unique(engagement_counts, return_counts=True)}")
        
        # Summary
        print("\n" + "="*80)
        print("📋 SUMMARY OF AUGUST METRICS")
        print("="*80)
        
        # Create summary data
        summary_data = {
            'Metric': [
                'Total Students in August',
                'Students with VWE Data',
                'Students with Industry Data',
                'Students with Session Data'
            ],
            'Count': [
                len(august_df),
                len(august_df['Virtual Work Experience'].dropna()) if 'Virtual Work Experience' in august_df.columns else 0,
                len(august_df['Industries'].dropna()) if 'Industries' in august_df.columns else 0,
                len(august_df['Web sessions'].dropna()) if 'Web sessions' in august_df.columns else 0
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        print(summary_df.to_string(index=False))
        
        return august_df
        
    except Exception as e:
        print(f"Error calculating metrics: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    august_df = calculate_august_metrics()
    if august_df is not None:
        print(f"\nAnalysis completed successfully!")
        print(f"Analyzed {len(august_df)} students in August")

import pandas as pd
import numpy as np

def demonstrate_excel_formula():
    """Demonstrate the Excel formula used for the Career Profiling engagement tracking"""
    
    print("EXCEL FORMULA FOR CAREER PROFILING ENGAGEMENT TRACKING")
    print("=" * 60)
    
    print("\n1. BASIC FORMULA (Column T):")
    print("=" * 40)
    print('=IF(ISNUMBER(SEARCH("Career Profiling Engaged",E2)),1,0)')
    
    print("\n2. ALTERNATIVE FORMULA (Column T):")
    print("=" * 40)
    print('=IF(COUNTIF(E2,"*Career Profiling Engaged*")>0,1,0)')
    
    print("\n3. MORE ROBUST FORMULA (Column T):")
    print("=" * 40)
    print('=IF(AND(E2<>"",ISNUMBER(SEARCH("Career Profiling Engaged",E2))),1,0)')
    
    print("\n4. FORMULA EXPLANATION:")
    print("=" * 40)
    print("• E2 = Person tag column (column E, row 2)")
    print("• SEARCH() = Finds text within a cell (case-insensitive)")
    print("• ISNUMBER() = Checks if SEARCH found a match")
    print("• IF() = Returns 1 if true, 0 if false")
    
    print("\n5. HOW TO APPLY IN EXCEL:")
    print("=" * 40)
    print("1. In cell T1, type: 'Career_Profiling_Flag' (header)")
    print("2. In cell T2, type the formula:")
    print("   =IF(ISNUMBER(SEARCH(\"Career Profiling Engaged\",E2)),1,0)")
    print("3. Copy cell T2 and paste down to all rows")
    print("4. Excel will automatically update E2 to E3, E4, etc.")
    
    print("\n6. PYTHON EQUIVALENT:")
    print("=" * 40)
    print("The Python code we used:")
    print("mask = august_df['Person tag'].str.contains('Career Profiling Engaged', na=False)")
    print("august_df.loc[mask, 'Career_Profiling_Flag'] = 1")
    
    print("\n7. EXAMPLE DATA:")
    print("=" * 40)
    
    # Create sample data to demonstrate
    sample_data = {
        'First name': ['Vito', 'Zizhu', 'Alaukika', 'John', 'Jane'],
        'Person tag': [
            '14 Engaged,Resume Builder Engaged,VWE Engaged',
            'Career Profiling Engaged,3 Engaged,28 Engaged',
            '28 Engaged,Career Profiling Engaged,Career Pr',
            'Video Profiling 5,Coaching Engaged',
            'Resume Builder Engaged,Skills Training'
        ]
    }
    
    df = pd.DataFrame(sample_data)
    df['Career_Profiling_Flag'] = df['Person tag'].str.contains('Career Profiling Engaged', na=False).astype(int)
    
    print(df)
    
    print("\n8. FORMULA RESULTS EXPLAINED:")
    print("=" * 40)
    for idx, row in df.iterrows():
        name = row['First name']
        person_tag = row['Person tag']
        flag = row['Career_Profiling_Flag']
        
        if flag == 1:
            print(f"✓ {name}: '{person_tag}' → Flag = 1 (Contains 'Career Profiling Engaged')")
        else:
            print(f"✗ {name}: '{person_tag}' → Flag = 0 (No 'Career Profiling Engaged')")

if __name__ == "__main__":
    demonstrate_excel_formula()

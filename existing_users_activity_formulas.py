import pandas as pd

def display_existing_users_activity_formulas():
    """Display Excel formulas for analyzing existing users' activity changes from July to August"""
    
    print("📊 EXCEL FORMULAS FOR EXISTING USERS ACTIVITY ANALYSIS")
    print("=" * 80)
    print("Analyzing how much activity of existing users from July was extended in August")
    print("=" * 80)
    
    print("\n🎯 1. IDENTIFYING EXISTING USERS")
    print("-" * 50)
    print("To find existing users (emails appearing in both July and August):")
    print("=IF(COUNTIF(August!B:B,July!B2)>0,\"Existing User\",\"New User\")")
    print("(Place this in a new column in July sheet to mark existing users)")
    
    print("\nAlternative approach using VLOOKUP:")
    print("=IF(ISNA(VLOOKUP(B2,August!B:B,1,FALSE)),\"New User\",\"Existing User\")")
    
    print("\n" + "="*80)
    print("📈 2. ACTIVITY COMPARISON FORMULAS")
    print("="*80)
    
    print("\n🔢 LOGIN COUNT / WEB SESSIONS COMPARISON")
    print("-" * 50)
    print("Formula: =August_Web_Sessions - July_Login_Count")
    print("Example: =VLOOKUP(B2,August!B:C,2,FALSE) - C2")
    print("Where:")
    print("• B2 = Email in July sheet")
    print("• August!B:C = Email and Web sessions columns in August")
    print("• C2 = Login Count in July sheet")
    print("• Positive result = Increased activity in August")
    print("• Negative result = Decreased activity in August")
    
    print("\n⏱️ AVERAGE LOGIN TIME COMPARISON (seconds)")
    print("-" * 50)
    print("Formula: =August_Avg_Login_Time - July_Avg_Login_Time")
    print("Example: =VLOOKUP(B2,August!B:D,3,FALSE) - D2")
    print("Where:")
    print("• B2 = Email in July sheet")
    print("• August!B:D = Email, Web sessions, and Avg Login Time columns")
    print("• D2 = Avg Login Time in July sheet")
    print("• Result in seconds")
    
    print("\n💼 VIRTUAL WORK EXPERIENCE (VWE) COMPARISON")
    print("-" * 50)
    print("Formula: =August_VWE - July_VWE")
    print("Example: =VLOOKUP(B2,August!B:K,10,FALSE) - K2")
    print("Where:")
    print("• B2 = Email in July sheet")
    print("• August!B:K = Email through VWE columns")
    print("• K2 = VWE in July sheet")
    
    print("\n" + "="*80)
    print("🔍 3. ADVANCED LOOKUP FORMULAS")
    print("="*80)
    
    print("\n📧 COMPREHENSIVE USER LOOKUP")
    print("-" * 40)
    print("For a complete user comparison table:")
    print("=VLOOKUP(B2,August!B:U,{2,3,4,10,19},FALSE)")
    print("This returns: Web sessions, Avg Login Time, Person tag, VWE, Career_Profiling_Flag")
    
    print("\n🔄 INDEX-MATCH ALTERNATIVE (More Flexible)")
    print("-" * 40)
    print("=INDEX(August!C:C,MATCH(B2,August!B:B,0))")
    print("This finds the row in August where email matches, then returns column C value")
    
    print("\n" + "="*80)
    print("📊 4. SUMMARY STATISTICS FORMULAS")
    print("="*80)
    
    print("\n👥 COUNTING EXISTING USERS")
    print("-" * 40)
    print("Total existing users:")
    print("=COUNTIF(July!Z:Z,\"Existing User\")")
    print("(Assuming column Z contains the existing user markers)")
    
    print("\nUsers with increased login activity:")
    print('=COUNTIF(AA:AA,">0")')
    print("(Assuming column AA contains login increase calculations)")
    
    print("\nUsers with increased login time:")
    print('=COUNTIF(AB:AB,">0")')
    print("(Assuming column AB contains time increase calculations)")
    
    print("\nUsers with increased VWE:")
    print('=COUNTIF(AC:AC,">0")')
    print("(Assuming column AC contains VWE increase calculations)")
    
    print("\n📊 AVERAGE CALCULATIONS")
    print("-" * 40)
    print("Average login increase:")
    print("=AVERAGE(AA:AA)")
    
    print("\nAverage time increase:")
    print("=AVERAGE(AB:AB)")
    
    print("\nAverage VWE increase:")
    print("=AVERAGE(AC:AC)")
    
    print("\n" + "="*80)
    print("🔧 5. CONDITIONAL ANALYSIS FORMULAS")
    print("="*80)
    
    print("\n🎯 MULTIPLE CRITERIA COUNTING")
    print("-" * 40)
    print("Users with both increased login AND increased time:")
    print('=COUNTIFS(AA:AA,">0",AB:AB,">0")')
    
    print("\nUsers with increased activity but decreased time:")
    print('=COUNTIFS(AA:AA,">0",AB:AB,"<0")')
    
    print("\nUsers with increased VWE and increased login:")
    print('=COUNTIFS(AA:AA,">0",AC:AC,">0")')
    
    print("\n" + "="*80)
    print("📋 6. IMPLEMENTATION GUIDE")
    print("="*80)
    
    print("\n📝 STEP-BY-STEP SETUP:")
    print("1. In July sheet, add new columns:")
    print("   • Column Z: 'Existing_User_Flag'")
    print("   • Column AA: 'Login_Increase'")
    print("   • Column AB: 'Time_Increase_Seconds'")
    print("   • Column AC: 'VWE_Increase'")
    print("   • Column AD: 'Activity_Summary'")
    
    print("\n2. In column Z (Existing_User_Flag), use:")
    print("=IF(COUNTIF(August!B:B,B2)>0,\"Existing User\",\"New User\")")
    
    print("\n3. In column AA (Login_Increase), use:")
    print("=IF(Z2=\"Existing User\",VLOOKUP(B2,August!B:C,2,FALSE)-C2,\"\")")
    
    print("\n4. In column AB (Time_Increase_Seconds), use:")
    print("=IF(Z2=\"Existing User\",VLOOKUP(B2,August!B:D,3,FALSE)-D2,\"\")")
    
    print("\n5. In column AC (VWE_Increase), use:")
    print("=IF(Z2=\"Existing User\",VLOOKUP(B2,August!B:K,10,FALSE)-K2,\"\")")
    
    print("\n6. In column AD (Activity_Summary), use:")
    print("=IF(Z2=\"Existing User\",CONCATENATE(\"Login: \",AA2,\", Time: \",AB2,\", VWE: \",AC2),\"\")")
    
    print("\n" + "="*80)
    print("⚠️ 7. IMPORTANT NOTES & TIPS")
    print("="*80)
    
    print("\n💡 FORMULA TIPS:")
    print("• Always use absolute references ($) when copying formulas across sheets")
    print("• Use IFERROR() to handle missing data: =IFERROR(VLOOKUP(...),\"No Match\")")
    print("• For percentage changes: =(August_Value - July_Value)/July_Value")
    print("• Use ROUND() for cleaner numbers: =ROUND(formula,2)")
    
    print("\n🚨 COMMON ISSUES:")
    print("• #N/A errors: Check if email exists in August sheet")
    print("• #VALUE! errors: Check for text in numeric columns")
    print("• #REF! errors: Check column references")
    
    print("\n✅ BEST PRACTICES:")
    print("• Test formulas on a few rows first")
    print("• Use named ranges for better readability")
    print("• Document your formulas in cell comments")
    print("• Backup your data before applying formulas")
    
    print("\n" + "="*80)
    print("🎯 8. QUICK REFERENCE FORMULAS")
    print("="*80)
    
    print("\n🔍 IDENTIFY EXISTING USER:")
    print("=IF(COUNTIF(August!B:B,B2)>0,\"Existing User\",\"New User\")")
    
    print("\n📊 CALCULATE LOGIN INCREASE:")
    print("=IF(Z2=\"Existing User\",VLOOKUP(B2,August!B:C,2,FALSE)-C2,\"\")")
    
    print("\n⏱️ CALCULATE TIME INCREASE:")
    print("=IF(Z2=\"Existing User\",VLOOKUP(B2,August!B:D,3,FALSE)-D2,\"\")")
    
    print("\n💼 CALCULATE VWE INCREASE:")
    print("=IF(Z2=\"Existing User\",VLOOKUP(B2,August!B:K,10,FALSE)-K2,\"\")")
    
    print("\n📈 COUNT INCREASED ACTIVITY:")
    print('=COUNTIF(AA:AA,">0")')
    
    print("\n📊 AVERAGE INCREASE:")
    print("=AVERAGE(AA:AA)")

if __name__ == "__main__":
    display_existing_users_activity_formulas()

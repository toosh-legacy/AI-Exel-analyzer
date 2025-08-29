# open_ai.py (Enhanced version)

import openai
import os
import pandas as pd
from typing import Dict, Any

# Load OpenAI key
openai.api_key = os.getenv("OPENAI_API_KEY") or "your-api-key-here"


def analyze_dataframe(df: pd.DataFrame, context: Dict[str, Any] = None) -> str:
    """
    Enhanced DataFrame analysis with better context and structured output
    """
    if context is None:
        context = {}

    # Generate data summary for the prompt
    data_summary = generate_data_summary(df)

    # Create context string
    context_str = ""
    if context:
        context_items = []
        if 'sheet_name' in context:
            context_items.append(f"Sheet: {context['sheet_name']}")
        if 'pivot_name' in context:
            context_items.append(f"Pivot Table: {context['pivot_name']}")
        if 'filters' in context:
            context_items.append(f"Applied Filters: {context['filters']}")
        context_str = " | ".join(context_items)

    prompt = f"""
    You are a financial analyst reviewing pivot table data from an Excel report.

    {f"Context: {context_str}" if context_str else ""}

    Data Summary:
    {data_summary}

    Raw Data (first 10 rows):
    {df.head(10).to_string(index=False)}

    Please provide a comprehensive financial analysis including:

    1. **Key Insights**: What are the main takeaways from this data?
    2. **Trends & Patterns**: Identify any notable trends, anomalies, or patterns
    3. **Financial Metrics**: Comment on important financial indicators present
    4. **Risk Assessment**: Highlight any potential risks or concerning areas
    5. **Recommendations**: Suggest actionable next steps or areas for further investigation
    6. **Data Quality**: Comment on data completeness and reliability

    Focus on actionable insights that would be valuable to management.
    Keep your analysis concise but thorough.
    """

    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[
                {
                    "role": "system",
                    "content": "You are an expert financial analyst with deep experience in corporate financial reporting, variance analysis, and business intelligence. Provide clear, actionable insights."
                },
                {"role": "user", "content": prompt}
            ],
            max_tokens=1500,
            temperature=0.3
        )
        return response['choices'][0]['message']['content'].strip()

    except Exception as e:
        return f"❌ OpenAI Analysis Error: {str(e)}\n\nFallback Analysis:\n{generate_fallback_analysis(df)}"


def generate_data_summary(df: pd.DataFrame) -> str:
    """Generate a structured summary of the DataFrame"""
    summary_parts = []

    # Basic info
    summary_parts.append(f"Shape: {df.shape[0]} rows × {df.shape[1]} columns")

    # Column info
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    text_cols = df.select_dtypes(include=['object', 'string']).columns.tolist()
    date_cols = df.select_dtypes(include=['datetime']).columns.tolist()

    if numeric_cols:
        summary_parts.append(
            f"Numeric columns ({len(numeric_cols)}): {', '.join(numeric_cols[:5])}{'...' if len(numeric_cols) > 5 else ''}")
    if text_cols:
        summary_parts.append(
            f"Text columns ({len(text_cols)}): {', '.join(text_cols[:5])}{'...' if len(text_cols) > 5 else ''}")
    if date_cols:
        summary_parts.append(f"Date columns ({len(date_cols)}): {', '.join(date_cols)}")

    # Data quality
    missing_data = df.isnull().sum()
    if missing_data.sum() > 0:
        cols_with_missing = missing_data[missing_data > 0]
        summary_parts.append(f"Missing data: {cols_with_missing.to_dict()}")

    # Numeric summaries
    if numeric_cols:
        summary_parts.append("\nNumeric column summaries:")
        for col in numeric_cols[:3]:  # First 3 numeric columns
            stats = df[col].describe()
            summary_parts.append(
                f"  {col}: min={stats['min']:.2f}, max={stats['max']:.2f}, mean={stats['mean']:.2f}, std={stats['std']:.2f}")

    return "\n".join(summary_parts)


def generate_fallback_analysis(df: pd.DataFrame) -> str:
    """Generate a basic analysis when OpenAI fails"""
    analysis_parts = []

    analysis_parts.append("BASIC DATA ANALYSIS:")
    analysis_parts.append(f"- Dataset contains {df.shape[0]} records across {df.shape[1]} fields")

    # Numeric analysis
    numeric_cols = df.select_dtypes(include=['number']).columns
    if len(numeric_cols) > 0:
        analysis_parts.append(f"\nNUMERIC ANALYSIS:")
        for col in numeric_cols:
            total = df[col].sum()
            avg = df[col].mean()
            analysis_parts.append(f"- {col}: Total = {total:,.2f}, Average = {avg:,.2f}")

    # Categorical analysis
    text_cols = df.select_dtypes(include=['object', 'string']).columns
    if len(text_cols) > 0:
        analysis_parts.append(f"\nCATEGORICAL ANALYSIS:")
        for col in text_cols[:3]:  # First 3 text columns
            unique_count = df[col].nunique()
            analysis_parts.append(f"- {col}: {unique_count} unique values")

    # Data quality
    missing_count = df.isnull().sum().sum()
    if missing_count > 0:
        analysis_parts.append(f"\nDATA QUALITY:")
        analysis_parts.append(f"- {missing_count} missing values detected")

    analysis_parts.append(f"\nRECOMMENDATION:")
    analysis_parts.append(f"- Review the data for completeness and accuracy")
    analysis_parts.append(f"- Consider trends and patterns in the numeric values")

    return "\n".join(analysis_parts)


def batch_analyze_dataframes(dataframes_dict: Dict[str, pd.DataFrame], context: Dict[str, Any] = None) -> Dict[
    str, str]:
    """
    Analyze multiple DataFrames in batch and return results
    """
    results = {}

    for name, df in dataframes_dict.items():
        print(f"🤖 Analyzing {name}...")

        # Add specific context for this dataframe
        df_context = context.copy() if context else {}
        df_context['pivot_name'] = name

        try:
            analysis = analyze_dataframe(df, df_context)
            results[name] = analysis
            print(f"✅ Analysis complete for {name}")
        except Exception as e:
            results[name] = f"❌ Analysis failed: {str(e)}"
            print(f"❌ Analysis failed for {name}: {e}")

    return results
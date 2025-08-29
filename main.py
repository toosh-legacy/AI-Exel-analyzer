# main.py (Cross-platform version)

# Import Excel and OpenAI utilities
from excel import get_unique_slicer_values, refresh_pivot_and_read, debug_excel_structure, analyze_excel_file_structure
from open_ai import analyze_dataframe, batch_analyze_dataframes
from itertools import product
import os
from datetime import datetime

# === Configuration ===
file_path = "CPRJune-25.xlsb"

# Multiple slicers supported per sheet
sheets_to_analyze = [
    {"sheet": "Monthly Variance Dynamic CPR", "slicers": ["Client", "VP", "Program"]},
    {"sheet": "CPR Common Size", "slicers": ["Client", "Type"]}
]


# Generate all possible slicer value combinations
def generate_slicer_combinations(slicer_values_map):
    keys = list(slicer_values_map.keys())
    values_product = list(product(*[slicer_values_map[k] for k in keys]))
    return [dict(zip(keys, combination)) for combination in values_product]


# === Analyze entire file structure first ===
print("Analyzing Excel file structure...")
analyze_excel_file_structure(file_path)

# === Debug specific sheets ===
print("\nAnalyzing target sheets...")
for config in sheets_to_analyze:
    debug_excel_structure(file_path, config["sheet"])

# === Process each sheet ===
for config in sheets_to_analyze:
    sheet = config["sheet"]
    slicer_fields = config["slicers"]

    print(f"\nProcessing sheet: {sheet} using slicers: {slicer_fields}")

    slicer_values_map = {}
    try:
        # Get all values for each slicer field
        for slicer in slicer_fields:
            slicer_values_map[slicer] = get_unique_slicer_values(file_path, sheet, slicer)
            print(f"  Values for '{slicer}': {len(slicer_values_map[slicer])} found")
    except Exception as e:
        print(f"Failed to get slicer values: {e}")
        continue

    # Generate every combination of slicer values (Cartesian product)
    slicer_combinations = generate_slicer_combinations(slicer_values_map)
    print(f"  Total combinations to process: {len(slicer_combinations)}")

    if len(slicer_combinations) > 20:
        print(f"  Large number of combinations detected!")
        response = input(f"  Continue with {len(slicer_combinations)} combinations? (y/n): ")
        if response.lower() != 'y':
            print(f"  Skipping sheet {sheet}")
            continue

    for combo_idx, combo in enumerate(slicer_combinations, 1):
        print(f"\n  Analyzing combination {combo_idx}/{len(slicer_combinations)}: {combo}")
        try:
            pivot_dataframes = refresh_pivot_and_read(file_path, sheet, combo)

            if not pivot_dataframes:
                print(f"    No data returned for combo {combo}")
                continue

            for pivot_name, df in pivot_dataframes.items():
                print(f"    Pivot Table: {pivot_name} - Shape: {df.shape}")

                if df.empty:
                    print(f"    Empty dataframe - skipping analysis")
                    continue

                # Check if DataFrame has meaningful data (more than just headers or single summary row)
                if len(df) >= 1 and len(df.columns) > 0:
                    print(f"    Running OpenAI analysis...")

                    try:
                        # Prepare context for OpenAI analysis
                        analysis_context = {
                            'sheet_name': sheet,
                            'pivot_name': pivot_name,
                            'filters': combo,
                            'combination_number': f"{combo_idx}/{len(slicer_combinations)}"
                        }

                        # Call enhanced OpenAI analysis
                        result = analyze_dataframe(df, analysis_context)

                        print(f"\n    OpenAI Analysis Results:")
                        print(f"    Sheet: {sheet} | Pivot: {pivot_name} | Filters: {combo}")
                        print("    " + "=" * 80)
                        print(result)
                        print("    " + "=" * 80)

                        # Save results to file
                        save_analysis_result(sheet, pivot_name, combo, result, df.shape)

                    except Exception as ai_error:
                        print(f"    OpenAI analysis failed: {ai_error}")
                else:
                    print(f"    Insufficient data for analysis (only {len(df)} rows)")

        except Exception as err:
            print(f"  Failed for slicer combo {combo}: {err}")


def save_analysis_result(sheet_name: str, pivot_name: str, combo: dict, analysis: str, data_shape: tuple):
    """Save analysis results to a file for later review"""
    try:
        # Create results directory if it doesn't exist
        results_dir = "analysis_results"
        os.makedirs(results_dir, exist_ok=True)

        # Create filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        combo_str = "_".join([f"{k}-{v}" for k, v in combo.items()])
        filename = f"{results_dir}/{sheet_name}_{pivot_name}_{combo_str}_{timestamp}.txt"

        with open(filename, 'w', encoding='utf-8') as f:
            f.write(f"Analysis Results\n")
            f.write(f"================\n")
            f.write(f"Timestamp: {datetime.now().isoformat()}\n")
            f.write(f"Sheet: {sheet_name}\n")
            f.write(f"Pivot Table: {pivot_name}\n")
            f.write(f"Filters Applied: {combo}\n")
            f.write(f"Data Shape: {data_shape[0]} rows × {data_shape[1]} columns\n")
            f.write(f"\nAnalysis:\n")
            f.write(f"---------\n")
            f.write(analysis)
            f.write(f"\n\nEnd of Analysis\n")

        print(f"    Analysis saved to: {filename}")

    except Exception as e:
        print(f"    Could not save analysis: {e}")
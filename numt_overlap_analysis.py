import pandas as pd
import matplotlib.pyplot as plt
import os

def analyze_numt_overlaps(file_path, query_start=10761, query_end=12137):
    """
    Analyze NUMT overlaps with a specified mitochondrial region.
    
    Parameters:
    file_path (str): Path to the Excel file containing NUMT data
    query_start (int): Start position of region of interest
    query_end (int): End position of region of interest
    
    Returns:
    tuple: (overlapping_numts DataFrame, summary statistics dictionary)
    """
    # Read the Excel file
    df = pd.read_excel(file_path)
    
    # Find overlapping regions
    overlapping = df[
        (df['Mt Start'] <= query_end) & 
        (df['Mt End'] >= query_start)
    ].copy()
    
    if not overlapping.empty:
        # Calculate overlap details for each matching NUMT
        overlapping['Overlap Start'] = overlapping['Mt Start'].apply(lambda x: max(x, query_start))
        overlapping['Overlap End'] = overlapping['Mt End'].apply(lambda x: min(x, query_end))
        overlapping['Overlap Length'] = overlapping['Overlap End'] - overlapping['Overlap Start']
        overlapping['Overlap Percentage'] = (overlapping['Overlap Length'] / 
                                           (query_end - query_start) * 100).round(2)
        
        # Categorize overlap types
        def categorize_overlap(row):
            if row['Mt Start'] <= query_start and row['Mt End'] >= query_end:
                return 'Complete'
            elif row['Mt Start'] <= query_start:
                return 'Partial (Left)'
            elif row['Mt End'] >= query_end:
                return 'Partial (Right)'
            else:
                return 'Internal'
        
        overlapping['Overlap Type'] = overlapping.apply(categorize_overlap, axis=1)
        
        # Calculate summary statistics
        summary_stats = {
            'total_overlaps': len(overlapping),
            'total_bases_covered': overlapping['Overlap Length'].sum(),
            'percent_query_covered': (overlapping['Overlap Length'].sum() / 
                                    (query_end - query_start) * 100).round(2),
            'max_overlap_length': overlapping['Overlap Length'].max(),
            'min_overlap_length': overlapping['Overlap Length'].min(),
            'mean_overlap_length': overlapping['Overlap Length'].mean().round(2)
        }
    else:
        summary_stats = {
            'total_overlaps': 0,
            'total_bases_covered': 0,
            'percent_query_covered': 0,
            'max_overlap_length': 0,
            'min_overlap_length': 0,
            'mean_overlap_length': 0
        }
    
    return overlapping, summary_stats

def visualize_overlaps(overlapping_df, query_start=10761, query_end=12137):
    """
    Create a visualization of the overlapping regions.
    
    Parameters:
    overlapping_df (pandas.DataFrame): DataFrame containing overlapping NUMTs
    query_start (int): Start position of region of interest
    query_end (int): End position of region of interest
    """
    if overlapping_df.empty:
        print("No overlaps found to visualize.")
        return
    
    fig, ax = plt.subplots(figsize=(12, 6))
    
    # Plot query region
    ax.plot([query_start, query_end], [0, 0], 'k-', linewidth=2, label='Query Region')
    
    # Create color map for overlap types
    color_map = {
        'Complete': 'green',
        'Partial (Left)': 'blue',
        'Partial (Right)': 'red',
        'Internal': 'purple'
    }
    
    # Plot overlapping NUMTs
    for idx, row in overlapping_df.iterrows():
        y_pos = idx + 1
        ax.plot([row['Mt Start'], row['Mt End']], [y_pos, y_pos], '-', 
                color=color_map[row['Overlap Type']], linewidth=2, 
                label=f"NUMT {row['NumtS Code']} ({row['Overlap Type']})")
        
    ax.set_xlabel('Mitochondrial Genome Position')
    ax.set_ylabel('NUMT Index')
    ax.set_title('NUMT Overlaps with Query Region')
    ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    ax.grid(True, linestyle='--', alpha=0.7)
    
    plt.tight_layout()
    return fig

def main():
    # Use relative paths
    script_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(script_dir, 'data')
    file_path = os.path.join(data_dir, '12864_2007_1460_MOESM1_ESM.xls')
    
    # Create output directory if it doesn't exist
    output_dir = os.path.join(script_dir, 'output')
    os.makedirs(output_dir, exist_ok=True)
    
    # Define query region
    query_start = 10761
    query_end = 12137
    
    # Analyze overlaps
    overlapping_numts, stats = analyze_numt_overlaps(file_path, query_start, query_end)
    
    # Print results
    print("\nOverlap Analysis Results:")
    print("-" * 50)
    print(f"Query Region: chrMT:{query_start}-{query_end}")
    print(f"Total length of query region: {query_end - query_start} bp")
    print("\nSummary Statistics:")
    for key, value in stats.items():
        print(f"{key.replace('_', ' ').title()}: {value}")
    
    if not overlapping_numts.empty:
        print("\nDetailed Overlapping NUMTs:")
        print(overlapping_numts[[
            'NumtS Code', 'Chr', 'Mt Start', 'Mt End', 
            'Overlap Start', 'Overlap End', 'Overlap Length', 
            'Overlap Percentage', 'Overlap Type'
        ]].to_string())
        
        # Create visualization
        fig = visualize_overlaps(overlapping_numts, query_start, query_end)
        
        # Save plot
        plot_path = os.path.join(output_dir, 'NUMT_overlap_visualization.png')
        plt.savefig(plot_path, bbox_inches='tight', dpi=300)
        plt.close()
        
        # Save results to Excel
        output_file = os.path.join(output_dir, 'NUMT_overlap_results.xlsx')
        overlapping_numts.to_excel(output_file, index=False)
        print(f"\nResults have been saved to {output_file}")
        print(f"Plot has been saved to {plot_path}")

if __name__ == "__main__":
    main()

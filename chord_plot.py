import pandas as pd
import holoviews as hv
from holoviews import opts
hv.extension('bokeh')
import os

# Get the list of files in the working directory
files = os.listdir()

# Filter files that start with "matrix_"
matrix_files = [file for file in files if file.startswith("matrix_")]

# Iterate over each matrix file
for matrix_file in matrix_files:
    # Read the matrix DataFrame from the file
    matrix_df = pd.read_excel(matrix_file, index_col=0)

    # Initialize an empty set to store chord combinations
    chord_combinations = set()
    chord_data = []

    # Iterate over rows and columns of the matrix DataFrame
    for source_index, source_row in matrix_df.iterrows():
        for target_index, value in source_row.items():
            # Skip if source and target are the same
            if source_index == target_index:
                continue
            # Skip if the combination is already in the chord combinations
            if (source_index, target_index) in chord_combinations or (target_index, source_index) in chord_combinations:
                continue
            # Append the source, target, and value to the chord data
            chord_data.append({'source': source_index, 'target': target_index, 'value': value})
            # Add the combination to the set of chord combinations
            chord_combinations.add((source_index, target_index))

    # Create the chord DataFrame from the chord data list
    chord_df = pd.DataFrame(chord_data)

    # Save the chord DataFrame to a CSV file
    chord_df.to_csv(f'chord_{matrix_file}.csv', index=False)

    # Create and save the Chord diagram
    chord_plot = hv.Chord(chord_df)
    # Customize the appearance of the Chord diagram
    chord_plot.opts(
        opts.Chord(
            cmap='glasbey_hv',  # Colormap for chords
            edge_cmap='glasbey_hv',  # Colormap for edges (choose the same single color)
            edge_color=hv.dim('source').str(),  # Edge color based on source
            labels='index',  # Label categories by index
            width=800,  # Width of the plot
            height=800,  # Height of the plot
            node_color=hv.dim('index').str(),  # Color of nodes based on index
            node_size=20,  # Size of nodes
            node_line_color='white',  # Color of node outlines
            node_line_width=2,  # Width of node outlines
            edge_line_color='source',  # Color of edges based on source
            edge_line_width=2,  # Width of edges
            edge_alpha=0.8,  # Opacity of the edges
            node_alpha=0.8,  # Opacity of the nodes
            label_text_font_size='12pt',  # Font size of category labels
            label_text_color='black',  # Color of category labels
            label_text_baseline='middle',  # Vertical alignment of labels
            label_text_align='left',  # Horizontal alignment of labels
        )
    )
    hv.save(chord_plot, f'chord_{matrix_file}.html')

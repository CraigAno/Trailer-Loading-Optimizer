import streamlit as st
import os
import tempfile
import shutil
from wall import run_trailer_optimizer
import pandas as pd

st.set_page_config(page_title="Trailer Loading Optimizer", layout="wide")

st.title("Westeel Trailer Loading Optimizer Web App")

st.markdown("""
Upload the BOM Excel file below to run the trailer loading optimization.
The app will automatically use the standard Solver and Roof Dimensions files.
""")

# Only BOM upload
country = st.text_input("Enter Destination Country (e.g., Canada or USA)", value="")
bom_file = st.file_uploader("Upload BOM Wall Parts Excel file", type=["xlsx"])

# Fixed paths for standard files
standard_solver_path = 'Solver (2).xlsx'
standard_roof_path = 'Roof Sheet Dimensions.xlsx'

# standard_solver_path = '/Users/craignyabvure/Desktop/Python Test Projects AGI/Trailer Loading/Solver (2).xlsx'
# standard_roof_path = '/Users/craignyabvure/Desktop/Python Test Projects AGI/Trailer Loading/Roof Sheet Dimensions.xlsx'

if bom_file:
    with st.spinner("Running trailer loading optimization..."):
        with tempfile.TemporaryDirectory() as tmpdirname:
            # Save BOM upload to temp
            bom_path = os.path.join(tmpdirname, "bom.xlsx")
            with open(bom_path, "wb") as f:
                f.write(bom_file.getbuffer())

            # Copy fixed solver and roof files into temp dir
            solver_path = os.path.join(tmpdirname, "solver.xlsx")
            roof_path = os.path.join(tmpdirname, "roof_dims.xlsx")
            shutil.copy(standard_solver_path, solver_path)
            shutil.copy(standard_roof_path, roof_path)

            # Run optimizer function
            summary_df, loading_plan_df, unloaded_df, visuals = run_trailer_optimizer(
                bom_path, solver_path, roof_path, country
            )

            st.success("Optimization complete!")

            # Show DataFrames in tabs
            tabs = st.tabs(["Summary", "Loading Plan", "Unloaded Items"])

            with tabs[0]:
                st.header("Summary")
                st.dataframe(summary_df)
            with tabs[1]:
                st.header("Loading Plan")
                st.dataframe(loading_plan_df)
            with tabs[2]:
                st.header("Unloaded Items")
                st.dataframe(unloaded_df)

            # Provide download for Excel output file
            excel_path = os.path.join(tmpdirname, "Loading_Plan_Output.xlsx")
            if os.path.exists(excel_path):
                with open(excel_path, "rb") as f:
                    st.download_button(
                        label="Download Excel Output",
                        data=f,
                        file_name="Loading_Plan_Output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            # Show generated visuals (PNGs only)
            st.header("Visualizations")
            for vis_path in visuals:
                if vis_path.endswith(".png"):
                    st.image(vis_path)
               # elif vis_path.endswith(".html"):
                   # with open(vis_path, "r") as f:
                        # html_str = f.read()
                    # st.components.v1.html(html_str width=1200, height=700, scrolling=True)
else:
    st.info("Please upload the BOM Excel file to run the optimizer.")

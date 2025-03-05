import sys
from win32com.client import DispatchEx
import os
import math

def round_up(value, step=1000):
    """Round up the value to the next multiple of `step`."""
    return math.ceil(value / step) * step

def adjust_max_value(max_value):
    """Custom logic to adjust max value based on specific conditions."""
    if 3000 <= max_value < 3500:
        return 3500  # If max value is between 3000 and 3500, set max to 3500
    return round_up(max_value, 1000)  # Otherwise, round up to the next multiple of 1000

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python python.py <filePath>")
        sys.exit(1)

    filePath = sys.argv[1]
    savePath = filePath.replace("/", "\\")
    base, ext = os.path.splitext(savePath)
    save_path = f"{base}_Modified{ext}"

    # Initialize Excel and open the workbook
    xl = DispatchEx("Excel.Application")
    xl.Visible = False  # Run Excel in the background
    wb = xl.Workbooks.Open(savePath)

    # Select the sheet named "Table"
    tableSheet = wb.Sheets("Table")
    tableSheet.Cells(1, 1).Value = ""
    tableSheet.Cells(2, 1).Value = "No."

    try:
        sheet = wb.Sheets(1)

        # Ensure there is at least one chart
        if sheet.ChartObjects().Count > 0:
            chart_object = sheet.ChartObjects(1)  # Reference to the chart object
            chart = chart_object.Chart
            chart.HasLegend = True  # Enable legend
            
            # Set Y-axis and X-axis titles
            chart.Axes(1).HasTitle = True  # Y-axis
            chart.Axes(1).AxisTitle.Text = "Time (s)"  # Y-axis title
            
            chart.Axes(2).HasTitle = True  # X-axis
            chart.Axes(2).AxisTitle.Text = "Bitrate (kbps)"  # X-axis title

            # Iterate over all series in the chart
            for series in chart.SeriesCollection():
                if series.Name == "Bitrate":
                    series.ChartType = 4  # xlLine (Line graph)
                    series.Format.Line.Weight = 1.5  # Set line weight
                    series.Format.Line.ForeColor.RGB = 255  # Blue color (RGB value)

                    # Get all the values in the series
                    values = series.Values
                    max_value = max(values)
                    
                    # Console the maximum value for debugging
                    print(f"Max Bitrate Value: {max_value}")

                    # Adjust max value based on custom conditions
                    y_axis_max = adjust_max_value(max_value)

                    # Set Y-axis maximum scale to the adjusted value
                    chart.Axes(2).MaximumScale = y_axis_max  # Axes(2) refers to the Y-axis (vertical axis)

                    # Set Y-axis to always start from 0
                    chart.Axes(2).MinimumScale = 0

                    # Calculate the average of the series
                    average_value = sum(values) / len(values)

                    # Calculate the standard deviation of the series
                    mean = average_value
                    variance = sum((x - mean) ** 2 for x in values) / len(values)
                    standard_deviation = math.sqrt(variance)

                    # Insert the average value text box under the chart
                    left_position = chart_object.Left
                    top_position = chart_object.Top + chart_object.Height + 10  # Positioning below the chart
                    width = 240
                    height = 30

            chart_shape_max = sheet.Shapes.AddTextbox(Orientation=1, Left=left_position, Top=top_position, Width=width, Height=height)
            chart_shape_max.TextFrame.Characters().Text = f"Max Bitrate: {max_value:.2f} kbps"
            chart_shape_max.TextFrame.Characters().Font.Bold = True  # Make text bold
            chart_shape_max.TextFrame.Characters().Font.Size = 11  # Set font size for clarity

            top_position += height + 5  # Position max bitrate text box below the standard deviation
            chart_shape_avg = sheet.Shapes.AddTextbox(Orientation=1, Left=left_position, Top=top_position, Width=width, Height=height)
            chart_shape_avg.TextFrame.Characters().Text = f"Average Bitrate: {average_value:.2f} kbps"
            chart_shape_avg.TextFrame.Characters().Font.Bold = True  # Make text bold
            chart_shape_avg.TextFrame.Characters().Font.Size = 11  # Set font size for clarity

            # Add text box for standard deviation
            top_position += height + 5  # Position standard deviation text box below the average
            chart_shape_std = sheet.Shapes.AddTextbox(Orientation=1, Left=left_position, Top=top_position, Width=width, Height=height)
            chart_shape_std.TextFrame.Characters().Text = f"Population Standard Deviation: {standard_deviation:.2f} kbps"
            chart_shape_std.TextFrame.Characters().Font.Bold = True  # Make text bold
            chart_shape_std.TextFrame.Characters().Font.Size = 11  # Set font size for clarity

        else:
            print("No charts found in the first sheet.")

        # Determine appropriate file format
        file_format = 51 if ext.lower() == ".xlsx" else 56  # 51 = .xlsx, 56 = .xls
        wb.SaveAs(save_path, FileFormat=file_format)

    except Exception as e:
        print(f"Error: {e}")

    finally:
        wb.Close(SaveChanges=1)
        xl.Quit()

    print(f"File saved as: {save_path}")

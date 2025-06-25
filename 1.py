import gspread
from gspread_dataframe import get_as_dataframe
import pandas as pd
import os
import plotly.graph_objects as go
import plotly.io as pio

# Force Plotly to use the browser renderer
pio.renderers.default = "browser"

def read_google_sheet():
    """Read data from Google Sheets and return as DataFrame"""
    
    # Check if service account file exists
    service_account_file = 'python-viz-demo-macro-lab-61731e10a604.json'
    if not os.path.exists(service_account_file):
        print(f"Error: Service account file '{service_account_file}' not found!")
        print("Please make sure the JSON file is in the same directory as this script.")
        return None
    
    try:
        # 使用服务账号认证
        gc = gspread.service_account(filename=service_account_file)
        
        # 打开表格（用 URL 或 ID 都行）
        sheet_url = 'https://docs.google.com/spreadsheets/d/1ZfKu09wgS1mjTcMDpVzWxyWvBcLkeJ3Ljn3Rp8ruRaA/edit?gid=422504657#gid=422504657'
        sheet = gc.open_by_url(sheet_url)
        
        # 选择工作表
        worksheet = sheet.worksheet("Data")
        
        # 读成 DataFrame
        df = get_as_dataframe(worksheet)
        
        return df
        
    except Exception as e:
        print(f"Error reading Google Sheet: {e}")
        return None

def create_stacked_bar_chart(df):
    """Create a stacked bar chart from the DataFrame"""

    year_col = df.columns[0]
    stack_cols = ['Lobbying', 'PAC', 'Super PAC', 'Dark money', 'Other outside']
    color_map = {
        'Lobbying': '#4F8DFD',      # blue
        'PAC': '#B6A6F7',           # purple
        'Super PAC': '#FFB6B6',     # pink
        'Dark money': '#B6F7F7',    # light cyan
        'Other outside': '#FFE066'  # yellow
    }
    # Ensure years are sorted and as integers for display
    df = df.sort_values(by=year_col)
    df[year_col] = df[year_col].astype(int)
    # Fill missing values with zero for all stack columns
    for col in stack_cols:
        if col not in df.columns:
            df[col] = 0
    df[stack_cols] = df[stack_cols].fillna(0)
    # Convert year to string for categorical axis
    df[year_col] = df[year_col].astype(str)
    fig = go.Figure()
    for col in stack_cols:
        fig.add_trace(go.Bar(
            name=col,
            x=df[year_col],
            y=df[col],
            marker_color=color_map[col],
            hovertemplate=f'{col}: %{{y}}<br>Year: %{{x}}<extra></extra>'
        ))
    fig.update_layout(
        title={
            'text': 'Total U.S. Lobbying and Election Spending, 1998-2018<br><span style="font-size:16px; color:gray">Data: OpenSecrets.org, based on Senate Office of Public Records</span>',
            'x': 0.5
        },
        xaxis_title='',
        yaxis_title='Billions $',
        barmode='stack',
        height=600,
        plot_bgcolor='white',
        legend=dict(
            title='',
            orientation='v',
            yanchor='top',
            y=1,
            xanchor='left',
            x=1.02,
            font=dict(size=14)
        ),
        font=dict(size=18, family='Arial', color='gray'),
        xaxis=dict(type='category', tickmode='array', tickvals=df[year_col].tolist()),
        yaxis=dict(range=[0, 8.1], dtick=2, showgrid=True, gridcolor='lightgray')
    )
    fig.update_xaxes(showgrid=False, tickangle=0)
    fig.update_yaxes(showgrid=True, gridcolor='lightgray')
    # Save the plot as HTML
    fig.write_html('stacked_bar_chart.html')
    return fig

if __name__ == "__main__":
    # Read the data
    df = read_google_sheet()
    
    if df is not None:
        print("DataFrame shape:", df.shape)
        print("\nFirst few rows:")
        print(df.head())
        
        print("\nDataFrame info:")
        print(df.info())
        
        print("\nColumn names:")
        print(df.columns.tolist())
        
        # Save to CSV for backup
        df.to_csv('google_sheet_data.csv', index=False)
        print("\nData saved to 'google_sheet_data.csv'")
        
        # Create and display the stacked bar chart
        fig = create_stacked_bar_chart(df)
        if fig:
            # Display the plot inline
            fig.show()
            print("\nStacked bar chart displayed!")
        else:
            print("Failed to create visualization.")
    else:
        print("Failed to read data from Google Sheets.")
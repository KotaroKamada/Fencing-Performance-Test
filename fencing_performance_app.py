import streamlit as st
import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# Plotlyが利用可能かチェック
try:
    import plotly.express as px
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False
    st.warning("Plotly library not found. Graph functionality will be disabled. Please add plotly to requirements.txt.")

# ページ設定
st.set_page_config(
    page_title="Fencing Performance Test",
    page_icon="🔲",
    layout="wide",
    initial_sidebar_state="expanded"
)

# カスタムCSS（シックなデザイン）
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #2D3748 0%, #1A202C 100%);
        padding: 2.5rem;
        border-radius: 0px;
        margin-bottom: 2rem;
        text-align: center;
        color: white;
        font-weight: 700;
        font-size: 2.8rem;
        box-shadow: 0 8px 32px rgba(45, 55, 72, 0.25);
        border-left: 6px solid #171923;
    }
    
    .section-header {
        background: linear-gradient(135deg, #4A5568 0%, #2D3748 100%);
        padding: 1.2rem 2rem;
        border-radius: 0px;
        color: white;
        font-weight: 600;
        margin: 2rem 0 1.5rem 0;
        font-size: 1.4rem;
        box-shadow: 0 4px 16px rgba(74, 85, 104, 0.2);
        border-left: 4px solid #1A202C;
    }
    
    .metric-card {
        background: linear-gradient(135deg, #718096 0%, #4A5568 100%);
        padding: 2rem;
        border-radius: 0px;
        margin: 0.75rem;
        color: white;
        text-align: center;
        box-shadow: 0 8px 24px rgba(113, 128, 150, 0.15);
        transition: all 0.3s ease;
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    
    .metric-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 32px rgba(113, 128, 150, 0.25);
    }
    
    .highlight-metric {
        font-size: 2.4rem;
        font-weight: 700;
        margin: 0.8rem 0;
        text-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    
    .metric-label {
        font-size: 1.2rem;
        margin-bottom: 0.8rem;
        opacity: 0.95;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .comparison-text {
        font-size: 1rem;
        opacity: 0.85;
        margin-top: 0.8rem;
        font-weight: 400;
    }
    
    .stDataFrame {
        background: white;
        border-radius: 8px;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.08);
        overflow: hidden;
    }
    
    .stDataFrame table {
        font-size: 1.1rem !important;
    }
    
    .stDataFrame th {
        background: linear-gradient(135deg, #F9FAFB 0%, #E5E7EB 100%) !important;
        color: #374151 !important;
        font-weight: 600 !important;
        font-size: 1.15rem !important;
        padding: 1rem !important;
        border: none !important;
    }
    
    .stDataFrame td {
        padding: 0.9rem !important;
        font-size: 1.1rem !important;
        border-bottom: 1px solid #E5E7EB !important;
    }
    
    .stDataFrame tr:hover {
        background-color: #F9FAFB !important;
    }
    
    .player-title {
        color: #2D3748;
        font-size: 2.2rem;
        font-weight: 700;
        margin-bottom: 1rem;
        padding: 1rem 0;
        border-bottom: 3px solid #718096;
    }
    
    .date-info {
        background: linear-gradient(135deg, #F7FAFC 0%, #E2E8F0 100%);
        padding: 1rem;
        border-radius: 8px;
        color: #2D3748;
        font-weight: 500;
        text-align: center;
        border: 1px solid #CBD5E0;
    }
    
    .graph-section {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        margin: 1.5rem 0;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.08);
        border: 1px solid rgba(0, 0, 0, 0.05);
    }
    
    .graph-title {
        color: #2D3748;
        font-size: 1.3rem;
        font-weight: 600;
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #E2E8F0;
    }
    
    .test-type-header {
        background: linear-gradient(135deg, #1A202C 0%, #2D3748 100%);
        padding: 1rem 2rem;
        border-radius: 8px;
        color: white;
        font-weight: 600;
        margin: 1.5rem 0;
        font-size: 1.2rem;
        box-shadow: 0 4px 16px rgba(26, 32, 44, 0.2);
        border-left: 4px solid #171923;
    }
</style>
""", unsafe_allow_html=True)

# データ読み込み関数
@st.cache_data
def load_data_from_file(uploaded_file):
    """アップロードされたファイルからデータを読み込む関数"""
    try:
        if uploaded_file.name.endswith('.xlsx') or uploaded_file.name.endswith('.xls'):
            df = pd.read_excel(uploaded_file)
        elif uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            st.error("Unsupported file format. Please upload Excel (.xlsx, .xls) or CSV file.")
            return pd.DataFrame()
        
        # Date列を日付型に変換
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'])
        
        return df
        
    except Exception as e:
        st.error(f"Data loading error: {str(e)}")
        return pd.DataFrame()

def get_test_config():
    """Test configuration"""
    return {
        'CMJ': {
            'name': 'Counter Movement Jump',
            'metrics': [
                'Jump Height',
                'Countermovement Depth', 
                'Braking RFD',
                'Avg. Braking Force',
                'Avg. Propulsive Force',
                'mRSI'
            ],
            'units': {
                'Jump Height': 'cm',
                'Countermovement Depth': 'm',
                'Braking RFD': 'N/s',
                'Avg. Braking Force': 'N',
                'Avg. Propulsive Force': 'N',
                'mRSI': ''
            },
            'highlight': ['Jump Height', 'mRSI', 'Avg. Propulsive Force'],
            'female_norms': {
                'Jump Height': {'mean': 33.65, 'std': 4.28},
                'mRSI': {'mean': 0.47, 'std': 0.08},
                'Braking RFD': {'mean': 6594.37, 'std': 1858.18}
            }
        },
        'IMTP': {
            'name': 'Isometric Mid-Thigh Pull',
            'metrics': [
                'Peak Force',
                'Relative Peak Force (BW)',
                'RFD 0-50 ms',
                'RFD 0-100 ms',
                'RFD 0-150 ms',
                'RFD 0-200 ms',
                'RFD 0-250 ms'
            ],
            'units': {
                'Peak Force': 'N',
                'Relative Peak Force (BW)': 'BW',
                'RFD 0-50 ms': 'N/s',
                'RFD 0-100 ms': 'N/s',
                'RFD 0-150 ms': 'N/s',
                'RFD 0-200 ms': 'N/s',
                'RFD 0-250 ms': 'N/s'
            },
            'highlight': ['Peak Force', 'Relative Peak Force (BW)', 'RFD 0-100 ms'],
            'female_norms': {
                'Relative Peak Force (BW)': {'mean': 42.45, 'std': 7.21},
                'RFD 0-250 ms': {'mean': 102.43, 'std': 23.89}
            }
        }
    }

def safe_get_best_value(data, column, default=None):
    """安全に最高値を取得する関数"""
    try:
        if column not in data.columns or data.empty:
            return default, default
        
        # null、NaN、空文字列、0を除外
        valid_data = data[data[column].notna()]
        valid_data = valid_data[valid_data[column] != '']
        valid_data = valid_data[valid_data[column] != 0]
        
        if valid_data.empty:
            return default, default
        
        # 数値に変換
        numeric_values = pd.to_numeric(valid_data[column], errors='coerce')
        clean_values = numeric_values.dropna()
        
        if clean_values.empty:
            return default, default
        
        # 最高値とその測定日を取得
        max_value = clean_values.max()
        max_idx = clean_values.idxmax()
        
        # 測定日を取得
        best_date = "N/A"
        if 'Date' in data.columns and max_idx in data.index:
            date_val = data.loc[max_idx, 'Date']
            if pd.notna(date_val):
                best_date = date_val.strftime('%Y-%m-%d')
        
        return float(max_value), best_date
        
    except Exception as e:
        return default, default
def safe_get_value(data, column, default=None):
    """安全に最新値を取得する関数"""
    try:
        if column not in data.columns or data.empty:
            return default
        
        # Exclude null, NaN, empty strings, and zeros
        valid_data = data[data[column].notna()]
        valid_data = valid_data[valid_data[column] != '']
        valid_data = valid_data[valid_data[column] != 0]
        
        if valid_data.empty:
            return default
        
        # Get latest valid data sorted by date
        if 'Date' in valid_data.columns:
            latest_valid = valid_data.sort_values('Date', ascending=False).iloc[0]
            value = latest_valid[column]
        else:
            value = valid_data.iloc[0][column]
        
        # Value validation
        if pd.isna(value) or value == '' or value == 0:
            return default
        
        # For numeric types
        if isinstance(value, (int, float, np.number)):
            if np.isfinite(value):
                return float(value)
        
        return default
        
    except Exception as e:
        return default

def safe_mean(series):
    """Safely calculate mean"""
    if series.empty:
        return None
    numeric_series = pd.to_numeric(series, errors='coerce')
    clean_series = numeric_series.dropna()
    clean_series = clean_series[clean_series != 0]
    return clean_series.mean() if len(clean_series) > 0 else None

def format_value(value, unit=""):
    """Safely format values"""
    if value is None or pd.isna(value):
        return "N/A"
    try:
        formatted_val = f"{float(value):.2f}"
        return f"{formatted_val}{unit}" if unit else formatted_val
    except:
        return "N/A"

def create_comparison_table(player_data, all_data, metrics, test_type, config):
    """Create comparison table"""
    table_data = []
    
    # Use only same test type data for average calculation
    test_data = all_data[all_data['Type'] == test_type]
    
    # Get female norms for this test type
    female_norms = config[test_type].get('female_norms', {})
    
    for metric in metrics:
        # Get latest player data
        player_val = safe_get_value(player_data, metric)
        
        # Get best player data
        best_val, best_date = safe_get_best_value(player_data, metric)
        
        # Calculate average of all athletes
        avg_val = safe_mean(test_data[metric])
        
        # Get female norm data
        female_norm_text = "N/A"
        if metric in female_norms:
            mean_val = female_norms[metric]['mean']
            std_val = female_norms[metric]['std']
            female_norm_text = f"{mean_val:.2f} ± {std_val:.2f}"
        
        # Get measurement date for latest value
        measurement_date = "N/A"
        if player_val is not None:
            valid_data = player_data.dropna(subset=[metric])
            valid_data = valid_data[valid_data[metric] != 0]
            if not valid_data.empty and 'Date' in valid_data.columns:
                latest_valid = valid_data.sort_values('Date', ascending=False).iloc[0]
                measurement_date = latest_valid['Date'].strftime('%Y-%m-%d') if pd.notna(latest_valid['Date']) else "N/A"
        
        # Format best value with date
        best_value_text = "N/A"
        if best_val is not None:
            best_value_text = f"{best_val:.2f}"
            if best_date != "N/A":
                best_value_text += f" ({best_date})"
        
        table_data.append({
            'Metric': metric,
            'Latest Value': format_value(player_val),
            'Test Date': measurement_date,
            'Personal Best': best_value_text,
            'Team Average': format_value(avg_val),
            'Female Fencer Norm': female_norm_text
        })
    
    return pd.DataFrame(table_data)

def create_trend_chart(player_data, metrics, title, units):
    """Create trend chart"""
    if not PLOTLY_AVAILABLE:
        return None
        
    if len(player_data) < 2:
        return None
    
    player_data = player_data.sort_values('Date')
    
    # Filter valid metrics
    available_metrics = []
    for metric in metrics:
        if metric in player_data.columns:
            data_with_values = player_data.dropna(subset=[metric])
            data_with_values = data_with_values[data_with_values[metric] != 0]
            if len(data_with_values) >= 2:
                available_metrics.append(metric)
    
    if not available_metrics:
        return None
    
    # Subplot configuration
    rows = (len(available_metrics) + 1) // 2
    cols = min(2, len(available_metrics))
    
    fig = make_subplots(
        rows=rows,
        cols=cols,
        subplot_titles=[f"<b>{metric}</b>" for metric in available_metrics],
        vertical_spacing=0.18,
        horizontal_spacing=0.15
    )
    
    colors = ['#2D3748', '#4A5568', '#718096', '#1A202C', '#A0AEC0', '#CBD5E0']
    
    for i, metric in enumerate(available_metrics):
        row = (i // 2) + 1
        col = (i % 2) + 1
        
        data_with_values = player_data.dropna(subset=[metric])
        data_with_values = data_with_values[data_with_values[metric] != 0]
        
        if len(data_with_values) >= 2:
            # Main trend
            fig.add_trace(
                go.Scatter(
                    x=data_with_values['Date'],
                    y=data_with_values[metric],
                    mode='lines+markers',
                    name=metric,
                    line=dict(
                        color=colors[i % len(colors)], 
                        width=4,
                        shape='spline',
                        smoothing=0.3
                    ),
                    marker=dict(
                        size=10, 
                        line=dict(width=3, color='white'),
                        symbol='circle'
                    ),
                    showlegend=False,
                    hovertemplate='<b>%{fullData.name}</b><br>Date: %{x}<br>Value: %{y:.2f}<extra></extra>'
                ),
                row=row, col=col
            )
            
            # Latest value annotation
            latest_point = data_with_values.iloc[-1]
            latest_value = latest_point[metric]
            
            display_text = f"{latest_value:.2f}"
            
            fig.add_annotation(
                x=latest_point['Date'],
                y=latest_value,
                text=display_text,
                showarrow=True,
                arrowhead=2,
                arrowsize=1,
                arrowwidth=2,
                arrowcolor=colors[i % len(colors)],
                bgcolor="white",
                bordercolor=colors[i % len(colors)],
                borderwidth=2,
                font=dict(size=11, color=colors[i % len(colors)]),
                row=row, col=col
            )
            
            unit = units.get(metric, '')
            fig.update_yaxes(
                title_text=f"{unit}" if unit else "Value",
                row=row, col=col,
                gridcolor='rgba(0,0,0,0.08)',
                linecolor='rgba(0,0,0,0.2)',
                title_font=dict(size=12, color='#2D3748'),
                tickfont=dict(size=10)
            )
            fig.update_xaxes(
                row=row, col=col,
                gridcolor='rgba(0,0,0,0.08)',
                linecolor='rgba(0,0,0,0.2)',
                tickfont=dict(size=10)
            )
    
    fig.update_layout(
        title=dict(
            text=title,
            x=0.5,
            font=dict(size=20, color='#2D3748', family='Arial Black')
        ),
        height=400 * rows,
        showlegend=False,
        plot_bgcolor='rgba(247, 250, 252, 0.3)',
        paper_bgcolor='white',
        margin=dict(l=50, r=50, t=80, b=50),
        font=dict(family="Arial")
    )
    
    return fig

def main():
    # Header
    st.markdown('<div class="main-header">Fencing Performance Test</div>', 
                unsafe_allow_html=True)
    
    # File upload
    uploaded_file = st.file_uploader(
        "Upload your performance data file",
        type=['xlsx', 'xls', 'csv'],
        help="Please upload Excel or CSV file containing performance test data"
    )
    
    if uploaded_file is None:
        st.info("Please upload a data file to begin analysis.")
        st.markdown("""
        ### Expected Data Format:
        The file should contain columns:
        - **Name**: Athlete name
        - **Date**: Test date
        - **Type**: Test type (CMJ or IMTP)
        - **CMJ metrics**: Jump Height, Countermovement Depth, Braking RFD, Avg. Braking Force, Avg. Propulsive Force, mRSI
        - **IMTP metrics**: Peak Force, Relative Peak Force (BW), RFD 0-50 ms, RFD 0-100 ms, RFD 0-150 ms, RFD 0-200 ms, RFD 0-250 ms
        """)
        st.stop()
    
    # Load data
    df = load_data_from_file(uploaded_file)
    if df.empty:
        st.error("Failed to load data.")
        st.stop()
    
    # Test configuration
    config = get_test_config()
    
    # Sidebar
    st.sidebar.header("Athlete Selection")
    
    # Athlete name selection
    available_names = df['Name'].dropna().unique()
    if len(available_names) == 0:
        st.error("No athlete data found.")
        st.stop()
    
    selected_name = st.sidebar.selectbox("Select Athlete", available_names)
    
    # Debug information in sidebar
    with st.sidebar.expander("Athlete Information"):
        st.write(f"Selected: {selected_name}")
        
        player_data = df[df['Name'] == selected_name]
        st.write(f"Total Tests: {len(player_data)}")
        
        cmj_count = len(player_data[player_data['Type'] == 'CMJ'])
        imtp_count = len(player_data[player_data['Type'] == 'IMTP'])
        st.write(f"CMJ: {cmj_count} tests, IMTP: {imtp_count} tests")
    
    # Get selected athlete data
    player_data = df[df['Name'] == selected_name]
    
    if player_data.empty:
        st.error(f"No data found for athlete '{selected_name}'.")
        return
    
    # Athlete information display
    col1, col2 = st.columns([3, 1])
    with col1:
        st.markdown(f'<div class="player-title">{selected_name}</div>', unsafe_allow_html=True)
    with col2:
        # Display data period
        all_dates = player_data['Date'].dropna().sort_values(ascending=False)
        if not all_dates.empty:
            latest_date = all_dates.iloc[0].strftime('%Y-%m-%d')
            oldest_date = all_dates.iloc[-1].strftime('%Y-%m-%d')
            st.markdown(f'<div class="date-info">Test Period: {oldest_date} ~ {latest_date}</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="date-info">Test Date: N/A</div>', unsafe_allow_html=True)
    
    # Process each test type
    for test_type, test_config in config.items():
        # Filter data for this test type
        test_player_data = player_data[player_data['Type'] == test_type]
        
        if test_player_data.empty:
            continue
        
        st.markdown(f'<div class="section-header">{test_config["name"]} ({test_type})</div>', unsafe_allow_html=True)
        
        # Key Indicators
        if test_config['highlight']:
            st.markdown("### Key Indicators")
            highlight_cols = st.columns(len(test_config['highlight']))
            
            for i, metric in enumerate(test_config['highlight']):
                with highlight_cols[i]:
                    player_val = safe_get_value(test_player_data, metric)
                    best_val, best_date = safe_get_best_value(test_player_data, metric)
                    # Average of all athletes for same test type
                    test_data = df[df['Type'] == test_type]
                    avg_val = safe_mean(test_data[metric])
                    unit = test_config['units'].get(metric, '')
                    
                    # Get female norm if available
                    female_norm_text = ""
                    if 'female_norms' in test_config and metric in test_config['female_norms']:
                        norm_data = test_config['female_norms'][metric]
                        female_norm_text = f"<br>Female Norm: {norm_data['mean']:.2f} ± {norm_data['std']:.2f}"
                    
                    # Personal best text
                    best_text = ""
                    if best_val is not None:
                        best_text = f"<br>Personal Best: {best_val:.2f}{unit}"
                        if best_date != "N/A":
                            best_text += f" ({best_date})"
                    
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">{metric}</div>
                        <div class="highlight-metric">{format_value(player_val, unit)}</div>
                        <div class="comparison-text">
                            Team Average: {format_value(avg_val, unit)}{best_text}{female_norm_text}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
        
        # Detailed data table
        st.markdown("### Detailed Data")
        available_metrics = [m for m in test_config['metrics'] if m in df.columns]
        
        if available_metrics:
            comparison_df = create_comparison_table(
                test_player_data, df, available_metrics, test_type, config
            )
            st.dataframe(comparison_df, use_container_width=True, hide_index=True)
            
            # Trend graph
            trend_fig = create_trend_chart(
                test_player_data, 
                available_metrics, 
                f"{test_config['name']} Progress", 
                test_config['units']
            )
            
            if trend_fig:
                st.markdown("### Progress Chart")
                st.plotly_chart(trend_fig, use_container_width=True, config={'displayModeBar': False})
            else:
                st.info("Progress chart requires at least 2 test sessions.")
        else:
            st.info(f"No {test_type} data available.")
    
    # Statistics
    with st.expander("Data Statistics"):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            total_athletes = len(df['Name'].unique())
            st.metric("Total Athletes", total_athletes)
        with col2:
            player_measurements = len(player_data)
            st.metric("Athlete's Tests", player_measurements)
        with col3:
            cmj_total = len(df[df['Type'] == 'CMJ'])
            st.metric("Total CMJ Tests", cmj_total)
        with col4:
            imtp_total = len(df[df['Type'] == 'IMTP'])
            st.metric("Total IMTP Tests", imtp_total)

if __name__ == "__main__":
    main()
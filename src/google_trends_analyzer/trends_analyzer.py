from pytrends.request import TrendReq
import pandas as pd
import numpy as np
from typing import List, Dict, Tuple
from datetime import datetime, timedelta
import os


class AITrendAnalyzer:
    def __init__(self, language: str = 'en-US', timezone: int = 360):
        """Initialize Google Trends connection."""
        self.pytrends = TrendReq(hl=language, tz=timezone)
        
    def calculate_trend_score(self, tool_name: str) -> Dict[str, float]:
        """
        Calculate a trend score for an AI tool using multiple indicators:
        1. Recent momentum (last 7 days vs previous 7 days)
        2. Overall growth trend (using linear regression)
        3. Volatility adjustment
        4. Peak analysis
        """
        # Get daily data for the past 90 days
        self.pytrends.build_payload([tool_name], timeframe='today 3-m')
        data = self.pytrends.interest_over_time()
        
        if data.empty:
            return {
                "trend_score": 0,
                "is_trending": False,
                "confidence": 0,
                "metrics": {}
            }

        # Remove isPartial column
        data = data.drop('isPartial', axis=1)
        
        # Calculate metrics
        metrics = {
            "momentum": self._calculate_momentum(data[tool_name]),
            "growth_trend": self._calculate_growth_trend(data[tool_name]),
            "volatility": self._calculate_volatility(data[tool_name]),
            "peak_ratio": self._calculate_peak_ratio(data[tool_name])
        }
        
        # Calculate final trend score
        trend_score = self._compute_final_score(metrics)
        
        # Determine if trending with confidence score
        is_trending, confidence = self._determine_trending_status(trend_score, metrics)
        
        return {
            "trend_score": round(trend_score, 2),
            "is_trending": is_trending,
            "confidence": round(confidence, 2),
            "metrics": {k: round(v, 2) for k, v in metrics.items()}
        }

    def _calculate_momentum(self, series: pd.Series) -> float:
        """Calculate recent momentum (last 7 days vs previous 7 days)."""
        recent = series[-7:].mean()
        previous = series[-14:-7].mean()
        return ((recent - previous) / (previous + 1)) * 100

    def _calculate_growth_trend(self, series: pd.Series) -> float:
        """Calculate overall growth trend using linear regression."""
        x = np.arange(len(series))
        y = series.values
        slope, _ = np.polyfit(x, y, 1)
        return slope * len(series)

    def _calculate_volatility(self, series: pd.Series) -> float:
        """Calculate volatility (adjusted standard deviation)."""
        return series.std() / series.mean() if series.mean() != 0 else 0

    def _calculate_peak_ratio(self, series: pd.Series) -> float:
        """Calculate ratio of current value to peak value."""
        peak = series.max()
        current = series.iloc[-1]
        return (current / peak) * 100 if peak != 0 else 0

    def _compute_final_score(self, metrics: Dict[str, float]) -> float:
        """
        Compute final trend score using weighted metrics:
        - Momentum: 20%
        - Growth trend: 35%
        - Peak ratio: 35%
        - Volatility penalty: 10%
        """
        weights = {
            "momentum": 0.2,
            "growth_trend": 0.35,
            "peak_ratio": 0.35,
            "volatility": -0.1  # negative weight as high volatility reduces reliability
        }
        
        score = sum(metrics[key] * weights[key] for key in weights.keys())
        return max(0, min(100, score))  # Normalize to 0-100

    def _determine_trending_status(self, 
                                 trend_score: float, 
                                 metrics: Dict[str, float]) -> Tuple[bool, float]:
        """
        Determine if the tool is trending and calculate confidence.
        
        Returns:
            Tuple of (is_trending: bool, confidence: float)
        """
        # Trending thresholds
        TREND_THRESHOLD = 30  # Minimum score to be considered trending
        
        # Confidence calculation based on metrics consistency
        confidence_factors = [
            metrics["momentum"] > 0,
            metrics["growth_trend"] > 0,
            metrics["peak_ratio"] > 50,
            metrics["volatility"] < 0.5
        ]
        
        confidence = (sum(confidence_factors) / len(confidence_factors)) * 100
        is_trending = trend_score >= TREND_THRESHOLD and confidence >= 60
        
        return is_trending, confidence
    
    def create_results_dataframe(self, results: Dict) -> pd.DataFrame:
        """
        Convert results dictionary to a structured DataFrame.
        """
        data = []
        for tool, result in results.items():
            row = {
                'Tool Name': tool,
                'Trend Score': result['trend_score'],
                'Is Trending': result['is_trending'],
                'Confidence': result['confidence']
            }
            # Add metrics
            for metric, value in result['metrics'].items():
                row[f'Metric_{metric}'] = value
            
            data.append(row)
        
        # Create DataFrame with ordered columns
        df = pd.DataFrame(data)
        column_order = [
            'Tool Name',
            'Trend Score',
            'Is Trending',
            'Confidence',
            'Metric_momentum',
            'Metric_growth_trend',
            'Metric_volatility',
            'Metric_peak_ratio'
        ]
        df = df[column_order]
        return df

    def export_to_excel(self, df: pd.DataFrame, filename: str = None) -> str:
        """
        Export DataFrame to Excel with formatting.
        """
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"ai_tools_trend_analysis_{timestamp}.xlsx"
        
        # Create Excel writer object
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        
        # Write DataFrame to Excel
        df.to_excel(writer, sheet_name='Trend Analysis', index=False)
        
        # Get workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Trend Analysis']
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D8E4BC',
            'border': 1
        })
        
        cell_format = workbook.add_format({
            'border': 1
        })
        
        percent_format = workbook.add_format({
            'border': 1,
            'num_format': '0.00%'
        })
        
        # Apply formats
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
        # Set column widths
        worksheet.set_column('A:A', 20)  # Tool Name
        worksheet.set_column('B:D', 15)  # Scores and Boolean
        worksheet.set_column('E:H', 18)  # Metrics
        
        # Add conditional formatting for Is Trending column
        worksheet.conditional_format(1, 2, len(df), 2, {
            'type': 'cell',
            'criteria': '=',
            'value': True,
            'format': workbook.add_format({'bg_color': '#C6EFCE'})
        })
        
        # Add conditional formatting for Trend Score
        worksheet.conditional_format(1, 1, len(df), 1, {
            'type': '3_color_scale',
            'min_color': "#FF9999",
            'mid_color': "#FFFF99",
            'max_color': "#99FF99"
        })
        
        writer.close()
        return filename


def main():
    analyzer = AITrendAnalyzer()
    
    # Example AI tools to analyze
    ai_tools = [
        "ChatGPT",
        "DALL-E",
        "Midjourney",
        "Claude AI",
        "Stable Diffusion",
        "Anthropic",
        "Bard AI",
        "Copilot"
    ]
    
    # Analyze each tool
    results = {}
    for tool in ai_tools:
        results[tool] = analyzer.calculate_trend_score(tool)
    
    # Create DataFrame
    df = analyzer.create_results_dataframe(results)
    
    # Print results to console
    print("\nAI Tools Trend Analysis:")
    print("-" * 50)
    print(df)
    
    # Export to Excel
    output_dir = "trend_analysis_results"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.join(output_dir, f"ai_tools_trend_analysis_{timestamp}.xlsx")
    
    exported_file = analyzer.export_to_excel(df, filename)
    print(f"\nResults exported to: {exported_file}")

if __name__ == "__main__":
    main()
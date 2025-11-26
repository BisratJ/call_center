# ALX Call Center Analytics Dashboard

A professional analytics dashboard for tracking and visualizing call center performance metrics at ALX Ethiopia.

## 📊 Overview

This interactive dashboard provides comprehensive insights into call center operations, helping teams monitor outreach effectiveness, analyze contact outcomes, and track engagement across multiple programs.

## ✨ Features

- **Real-time Metrics**: Track total calls, reach rates, unique contacts, and engagement levels
- **Monthly Performance**: Visualize trends and compare performance across different time periods
- **Call Outcomes Analysis**: Detailed breakdown of reached, not reached, busy, and other call statuses
- **Interactive Charts**: Dynamic visualizations using Chart.js with data labels
- **Campaign Tracking**: Monitor performance across different campaign types
- **Key Insights**: Automated insights highlighting top performers and areas for improvement
- **Responsive Design**: Optimized for desktop and mobile viewing

## 🛠️ Technology Stack

- **Frontend**: Pure HTML/CSS/JavaScript
- **Visualization**: Chart.js v4.4.0 with datalabels plugin
- **Typography**: Inter font family
- **Styling**: Custom CSS with CSS variables
- **Hosting**: Ready for deployment on Vercel or any static hosting

## 📁 Project Structure

```
call_center/
├── index.html          # Main dashboard file
├── README.md           # Project documentation
├── vercel.json         # Deployment configuration
└── .gitignore          # Git ignore rules
```

## 🚀 Getting Started

1. Clone the repository:

   ```bash
   git clone https://github.com/BisratJ/call_center.git
   ```

2. Open `index.html` in your browser or serve using a local server:

   ```bash
   python -m http.server 8080
   ```

3. Navigate to `http://localhost:8080` to view the dashboard

## 📊 Dashboard Sections

- **Hero Section**: Overview metrics with key performance indicators
- **Executive Snapshot**: Quick stats on contacts reached, engagement, and top performers
- **Visual Analytics**: Interactive charts showing monthly trends and distributions
- **Performance Matrix**: Detailed monthly breakdown with reach rates
- **Call Outcomes**: Analysis of all call statuses and outcomes
- **Key Insights**: Data-driven insights and recommendations
- **Monthly Timeline**: Narrative summary of each month's performance

## 🎨 Customization

The dashboard uses CSS variables for easy theming. Modify the `:root` section in `index.html` to customize colors:

```css
:root {
  --primary: #2563eb;
  --primary-light: #60a5fa;
  --secondary: #10b981;
  /* ... other variables */
}
```

## 📱 Responsive Design

The dashboard is fully responsive and adapts to different screen sizes:

- Desktop: Full multi-column layout
- Tablet: Adaptive grid system
- Mobile: Single column with optimized spacing

## 📄 License

© 2025 ALX Ethiopia Call Center Analytics

---

**Built with ❤️ for data-driven decision making**

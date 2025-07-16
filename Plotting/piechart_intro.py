import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.font_manager as fm
import os

import numpy as np

# === Font setup ===
font_path = 'C:/Windows/Fonts/OpenSans-Regular.ttf'  # Update if needed
if os.path.exists(font_path):
    open_sans = fm.FontProperties(fname=font_path)
    plt.rcParams['font.family'] = open_sans.get_name()
else:
    print("Open Sans font not found, using default.")
    plt.rcParams['font.family'] = 'sans-serif'

# === Emission data ===
global_emissions = 53.0
industry_emissions = 11.247
chemical_emissions = 5.6
hvc_emissions = 0.7 * chemical_emissions

# # === Main pie chart (Global emissions) ===
# other_emissions = global_emissions - industry_emissions
# main_labels = ['Industry', 'Other sectors']
# main_sizes = [industry_emissions, other_emissions]
# main_colors = ['#fdae61', '#d0d0d0']
# main_explode = (0.1, 0)  # Pop out industry

# === Main pie chart type 2 (Global emissions) ===
other_emissions = global_emissions - industry_emissions
other_industry = industry_emissions - chemical_emissions
other_chemical = chemical_emissions - hvc_emissions
main_labels = ['Other sectors\n(non industry)', 'Other industry', 'Other chemical industry', 'Fertilizer-olefin\nproduction']
main_sizes = [other_emissions, other_industry, other_chemical, hvc_emissions]
main_colors = ['#d0d0d0', '#804E49', '#E7DECD', '#B5A27D']
# main_colors = ['#d0d0d0', '#fdae61', '#fee08b', '#d73027']
main_explode = (0, 0.2, 0.2, 0.3)  # Pop out industry

# === Sub pie chart (Industry breakdown) ===
# other_industry = industry_emissions - chemical_emissions
# other_chemical = chemical_emissions - hvc_emissions
# industry_labels = ['Other industry', 'Other chemical industry', 'HVC production']
# industry_sizes = [other_industry, other_chemical, hvc_emissions]
# industry_colors = ['#fdae61', '#fee08b', '#d73027']
# industry_explode = (0.05, 0.05, 0.05)

# === Create the figure ===
fig, ax = plt.subplots(figsize=(8, 7))
ax.axis('equal')  # Keeps circles round

# --- Main Pie ---
wedges, texts, autotexts = ax.pie(
    main_sizes,
    colors=main_colors,
    labels=main_labels,
    startangle=90,
    explode=main_explode,
    shadow=True,
    autopct='%1.0f%%',
    pctdistance=0.8,
    labeldistance=1.1,
    wedgeprops={'edgecolor': 'none'}
)


# 'texts' is returned from ax.pie()
for text in texts:
    text.set_fontproperties(open_sans)
    text.set_fontsize(14)
    text.set_weight('bold')
    text.set_color('#262626')

for autotext in autotexts:
    autotext.set_fontproperties(open_sans)
    autotext.set_fontsize(12)
    autotext.set_weight('bold')
    autotext.set_color('#111211')

# # === Manually add labels ===
# label_distances = [1.4, 1.5, 1.6, 1.5]
# for i, wedge in enumerate(wedges):
#     angle = (wedge.theta2 + wedge.theta1) / 2
#     x = label_distances[i] * np.cos(np.radians(angle))
#     y = label_distances[i] * np.sin(np.radians(angle))
#
#     ax.text(
#         x, y,
#         main_labels[i],
#         ha='center',
#         va='center',
#         fontproperties=open_sans,
#         fontsize=12
#     )



# === Title and save ===
# plt.suptitle('Global Greenhouse Gas Emissions (2023)', fontsize=16, weight='bold')
savepath = 'C:/Users/5637635/Documents/OneDrive - Universiteit Utrecht/MyPhD/Thesis/Thesis_plots/'
plt.savefig(f"{savepath}emissions_piechart.pdf", format='pdf', bbox_inches='tight')

plt.show()

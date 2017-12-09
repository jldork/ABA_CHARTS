import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.patches as mpatches
import matplotlib.pyplot as plt
import matplotlib.lines as mlines
import pandas as pd
from sklearn import linear_model

def graph(excel_file, sheet_name):
    #############################
    # Reading and Cleaning Data #
    #############################

    sheet = pd.read_excel(excel_file, sheet_name=sheet_name)
    data = sheet[['Objective Number','Protocol','Date Met','Lesson LU','Objectives Met ', 'Tactic']]
    data = data[~data['Lesson LU'].isnull()] # Filter to latest LU date
    data['Objective Number'] = data['Objective Number'].astype(int) 
    data['Cumulative Objectives'] = data['Objectives Met '].cumsum()
    data.fillna(0, inplace=True)

    ###################
    # Build Trendline #
    ###################

    # Create linear regression
    regr = linear_model.LinearRegression()
    regr.fit(data[['Objective Number']], data[['Cumulative Objectives']])
    predicted = regr.predict(data[['Objective Number']])
    data['Trendline'] = predicted[:,0]
    
    ###################
    # Build the Chart #
    ###################

    plot_LU = data.plot(x='Objective Number', y='Lesson LU', kind='bar', color='black', figsize=(18,11), legend=False)
    plot_Tactic = data.plot(x='Objective Number', y='Tactic', kind='bar', color='red', bottom=data['Lesson LU'], ax=plot_LU, legend=False)
    plot_Tactic = data.plot(x='Objective Number', y='Protocol', kind='bar', color='blue', bottom=data['Protocol'], ax=plot_LU, legend=False)
    plot_LU.set_ylabel('Learn Units to Meet an Objective')
    
    ax = data.plot(x='Objective Number', y='Cumulative Objectives', secondary_y=True, ax=plot_Tactic, color='g', legend=False)
    ax2 = ax.twiny()
    label = ax.set_ylabel('Cumulative Number of Objectives Met', labelpad=15 )
    label.set_rotation(270)
    dates = data['Date Met'].dt.strftime('%d-%b')
    ax2.set_xticks(data['Objective Number'])
    ax2.set_xticklabels(dates)
    for xlabel_i in ax2.get_xticklabels()[0:len(data):2]:
        xlabel_i.set_visible(False)
    plt.title(sheet_name + ' - 1 - Math', y=1.08)
    ax = data.plot(x='Objective Number', y='Trendline', secondary_y=True, ax=plot_Tactic, color='black', linestyle='--', alpha=0.5, legend=False)
    
    ##########
    # LEGEND #
    ##########

    # Block Legend Markers
    red_patch = mpatches.Patch(color='red', label='Tactic')
    black_patch = mpatches.Patch(color='black', label='Lesson LU')
    blue_patch = mpatches.Patch(color='blue', label='Protocol')
    
    # Linear Legend Markers
    green_patch = mlines.Line2D([], [], color='green', markersize=5, label='Cumulative Objectives')
    trendline = mlines.Line2D([], [], color='black', label='Linear (Cumulative Objectives)')
    
    # Add Legend
    plt.legend(
        bbox_to_anchor=(0.19, 1.01,.6, 0.075), 
        handles=[red_patch, black_patch, blue_patch, green_patch, trendline], 
        mode='expand', ncol=4)
    
    ########################
    # Saving and Uploading #
    ########################

    # Save Figure to /tmp
    fig = plt.gcf()
    chart_fname = '/tmp/' + student + '.png'
    fig.savefig(chart_fname)

    return open(chart_fname, 'wb')

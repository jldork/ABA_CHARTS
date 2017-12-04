import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.patches as mpatches
import matplotlib.pyplot as plt
import pandas as pd
from sklearn import linear_model

student_sheets = pd.read_excel('./LPEverydayMath.xlsx', sheet_name=None)
students = ['Ashley', 'Benny','Luisana', 'Declan', 'Grayson', 'Vera']
print("Compiling graphs for: \n", ', '.join(students))

def graph(sheets, sheet_name):
    sheet = sheets[sheet_name]
    data = sheet[['Objective Number','Date Met','Lesson LU','Objectives Met ', 'Tactic']]
    data = data[~data['Lesson LU'].isnull()] # Filter to latest LU date
    data['Objective Number'] = data['Objective Number'].astype(int) 
    data['Cumulative Objectives'] = data['Objectives Met '].cumsum()
    data.fillna(0, inplace=True)
    

    # Create linear regression object
    regr = linear_model.LinearRegression()
    # Train the model using the training sets
    regr.fit(data[['Objective Number']], data[['Cumulative Objectives']])
    # Make predictions using the testing set
    predicted = regr.predict(data[['Objective Number']])
    data['Trendline'] = predicted[:,0]
    
    plot_LU = data.plot(x='Objective Number', y='Lesson LU', kind='bar', color='black', figsize=(18,11), legend=False)
    plot_Tactic = data.plot(x='Objective Number', y='Tactic', kind='bar', color='red', bottom=data['Lesson LU'], ax=plot_LU, legend=False)
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
    
    red_patch = mpatches.Patch(color='red', label='Tactic')
    black_patch = mpatches.Patch(color='black', label='Lesson LU')
    green_patch = mpatches.Patch(color='green', label='Cumulative Objectives')
    trendline = mpatches.Patch(color='black', linestyle='dashed', label='Linear (Cumulative Objectives)')
    plt.legend(
        bbox_to_anchor=(0.19, 1.01,.6, 0.075), 
        handles=[red_patch, black_patch, green_patch, trendline], 
        mode='expand', ncol=4)
    
    return ax

if not os.path.exists('./charts'):
    os.makedirs('./charts')

for student in students:
    ax = graph(student_sheets,student)
    fig = plt.gcf()
    fig.savefig('charts/' + student + '.png')

print("Complete! Check the charts folder")

# Data Science Job Market Analysis and Excel Dashboard

![final created dashboard]()

## Introduction

As an aspiring data analyst, this project was created with the aim of practicing and demonstrating proficiency in advanced Excel skills while analyzing and gaining valuable insights into the job market of data science jobs.

The dataset used for this project contains over 30,000 real-world data science job postings collected from 2023 and was obtained from the [Excel Data Analytics Course](https://github.com/lukebarousse/Excel_Data_Analytics_Course) by Luke Barousse.

After finishing the course, I realized that I wanted to perform more in-depth analysis than what was presented in the final projects within the course. Using the skills I learned, I took it upon myself to further clean and transform the dataset, to perform and visualize deeper analysis, and to build a more customizable overview of important data within an interactive dashboard created inside Excel.

What can be learned about the job market by diving into different variables and visualizing their relationships? To guide my data exploration, I sought to answer the following questions:

1. How do yearly and hourly salaries compare with each other? 
2. What are the most valuable skills to know?
3. How does the frequency of job postings vary throughout the year?
4. How many job postings are located remotely and how does this proportion vary throughout the year?
5. How do different schedule types make up total job postings?

For each of these questions, I wanted to be able to compare the answers when including all job postings or when narrowed down to job titles of choice for tailored insight. As I am personally most interested in data analyst roles, my analysis will especially focus on these postings.

### Excel Skills Used in Analysis (ETL, EDA, & Visualization):

* Power Query
* Power Pivot/data models
* PivotTable/PivotChart
* DAX (Data Analysis Expressions)
* Functions (IF, FILTER, XLOOKUP, SORT, UNIQUE)
* Dynamic arrays
* Data validation
* Conditional formatting

## ETL (Extract, Transform, Load)

First, I used Power Query and M language to extract the data from the Excel file `data_jobs.xlsx` provided by the course into a new workbook as well as to add two new columns: an index of job IDs and a multiplication of hourly salaries by working hours in a year for easier comparison against yearly salaries.

This query, seen below, contains messy lists of skills and multiple schedule options within single rows, so I referenced this query to create two additional queries in order to clean and make these portions of data usable in future PivotTables.

![data_jobs query]()

In the query `data_jobs_skills`, I removed all columns except job IDs and job skills, removed extra characters, split and expanded the lists of skills to contain one skill per row, and removed existing duplicates.

![data_jobs_skills query]()

In the query `data_jobs_schedules`, I again narrowed down to two columns and expanded lists of schedule types corresponding to each job ID. 

![data_jobs_schedules query]()

Finally, I loaded all three queries to the data model within Power Pivot and defined their relationships using corresponding job IDs.

![power pivot data model relationships]()

## EDA (Exploratory Data Analysis) & Visualization

In order to explore relationships, find answers to my questions, and build a dashboard using PivotTables, I would need to use Power Pivot to create measures, or custom calculations more advanced than default PivotTable functions.

### 1. How do yearly and hourly salaries compare with each other?

To answer this question, I created a measure using DAX functions to calculate the median yearly salary:
```
=MEDIAN(data_jobs[salary_year_avg])
```
Identical measures were also created for hourly and adjusted hourly salary columns and all placed in a PivotTable under Values with `job_title_short` column under Rows. I sorted the table by yearly salary and applied conditional formatting to the yearly and adjusted hourly salaries together in order to highlight differences between values.

These results were then visualized in a PivotChart combining clustered columns for the yearly and adjusted hourly salaries with markers for original hourly salaries on a secondary axis. I adjusted axis bounds so that the original and adjusted hourly salaries would line up.

I also was curious about the distribution of yearly and hourly salaries for different job titles. However, histograms cannot be used with PivotTables. Instead, I created two more queries in Power Query narrowing columns down to job titles and either yearly or hourly salaries, removing null values, and loaded these tables into a new worksheet. Next, I created a dynamic array sorting all unique job titles:

```
=SORT(UNIQUE(data_jobs_yearly_salaries[job_title_short]))
```

I also wrote "All Job Titles" in the cell above the array. Then I used this list to define a drop-down selection of job titles using data validation. Next, I wrote formulas to create dynamic arrays of only the salaries where the job title is currently selected in the drop-down cell. I used these dynamic arrays to create my histograms:

```
=IF(
    'Salary Analysis'!$G$15="All Job Titles",
    data_jobs_yearly_salaries[salary_year_avg],
    FILTER(
        data_jobs_yearly_salaries[salary_year_avg],
        data_jobs_yearly_salaries[job_title_short]=XLOOKUP(
            'Salary Analysis'!$G$15,
            data_jobs_yearly_salaries[job_title_short],
            data_jobs_yearly_salaries[job_title_short])
))
```
### Analysis

![Salary Analysis worksheet containing combo chart and histograms]()

* Looking at the Yearly vs. Hourly Salaries Comparison chart, data engineers, data scientists, and senior roles recieve the greatest yearly salaries with business analysts having the lowest yearly salary. Data engineers and senior data engineers have the highest hourly salaries with data analyst having the lowest hourly salary.
* Across every job title, the median yearly salary is greater than a year's worth of hourly salary for the same title, with business analyst yearly salary being barely higher than hourly and senior data scientist yearly salary being over twice as much as hourly salary for the same title. This suggests it might be easier to find higher pay when looking at jobs offering yearly salaries, especially for the titles displaying larger differences.
* When including all job titles, the histograms reveal a very wide range of salaries. Many yearly salaries fall between $82K-$121K, peaking between $108K-$121K, and drop off before $68K and after $161K. The hourly salaries show the highest cluster between $19-$30 with another cluster around $47-$58. Both yearly and hourly distributions extend towards the upper end indicating significantly higher paying outliers.

#### What About Data Analysts?

![Histograms filtered to data analyst]()

* Many yearly salaries fall between $70K-$106K, peaking between $88K-$97K, and drop off before $52K and after $115K. Hourly salaries have a high spike between $22-$26. Both yearly and hourly distributions again extend at the upper end, indicating significantly higher-paying jobs exist but are relatively few.

### 2. What are the most valuable skills to know?

To answer this question, I wanted to find the skills with the highest demand as well as being associated with higher pay. By placing a count on `job_id` grouping by `job_skills` in a PivotTable, and limiting to top ten while sorting by highest counts, I found the top ten most frequently mentioned skills. I also added a job title slicer to update the top skills to different roles and applied conditional formatting to highlight the jumps in skill frequencies. I then created a bar chart to compare these skills by job posting counts.

In order to see the salaries associated with these top skills, I created an indentical PivotTable and added median yearly salaries also being grouped by job skills. I then sorted by salary and created a chart with columns for the salaries and markers displaying job counts per skill.

### Analysis

![Job Skills Analysis worksheet containing top skills and skill salaries]()

* When including all job titles, the charts as well as conditional formatting on the tables reveals a large jump between the number of postings mentioning SQL and Python, 18,000 and 17,689  respectively, and the next highest demand skill Tableau, mentioned in 7,043 postings.
* Spark is the skill out of the top ten most frequent skills associated with the highest median salary, $140,000, and is mentioned in 5,290 job postings.

#### What About Data Analysts?

![Job Skills Analysis filtered to data analyst]()

* For data analyst jobs, SQL remains the skill in highest demand at 5,033 postings followed by Excel at 3,839. Python has dropped to third highest demand skill at 2,765 with Tableau close behind.
* Of the ten highest demand skills, Python is associated with the highest salaries with a median yearly salary of $98,500. Almost all median salaries tied to these skills are $90K or more.
* SQL, Python, and Tableau appear to be valuable skills to know as a data analyst as they are all in high demand and associated with hiher salaries.

### 3. How does the frequency of job postings vary throughout the year?

To find the answer, I created a PivotTable of job counts by `job_posted_date`, which automatically group into months of the year as rows. I again added condtional formatting, a slicer to filter by job titles, and made a line chart.

### Analysis

![Monthly job postings chart]()

* When including all data science jobs, job postings rise and fall, peaking in August at 3554 postings, then decline towards the end of the year, with November having the fewest postings of 1,963. Thus, it might generally be easier to find job opportunities earlier in the year.

#### What About Data Analysts?

* Filtering to only data analyst job postings, postings are high in January with 982 postings, peak in August at 1031 postings, and then fall in later months with the lowest posting count of 645 in November, again suggesting more data analyst job opportunities in the earlier months of the year.

### 4. How many job postings are located remotely and how does this proportion vary throughout the year?

I adapted the previous chart containing total job postings and created the measure `Remote Job Count` to count only remote job postings:
```
=CALCULATE(COUNTROWS(data_jobs), data_jobs[job_work_from_home]=TRUE)
```
I also created a measure referencing the previous one to display the percentage of remote job postings from total job postings:
```
=[Remote Job Count]/COUNT(data_jobs[job_id])
```
I added both of these measures to the PivotTable values, added conditional formatting to each column, and created a line chart with markers for remote job percentages on a secondary axis.

### Analysis

![Postings Analysis worksheet containing monthly remote postings line chart]()

* When including all data science jobs, remote postings rise and fall in proportion with total postings as seen by the steady percentages in the first half of the year. October saw the highest count of remote postings at 612 while November saw the lowest count of 371, December barely having more postings than that. This again indicates jobs being in lower supply at the end of the year.
* Remote job postings remained relatively steady in the months that total jobs peaked, dropping the remote job percentage to 12.7% in August. Remote postings then rose while total postings decreased, causing October to see the highest percentage of remote job postings at 24.0%. One possible reason for the sudden increase of in-person jobs after May and the fluctuations of remote jobs may have been the world adjusting to the ending of the COVID-19 pandemic in 2023.

#### What About Data Analysts?

* For data analysts, remote jobs are high in January at 114 postings, then fewest from February to April at 78-79, and peaking in October at 137. The lowest remote job percentage was in March at 10.2% with another dip during August as in-person jobs peaked. The following months saw total jobs drop while remote jobs peaked, with the highest percentage of remote jobs being 17.2% in October.
* The percentage of remote jobs for data analysts is lower than the percentage for all data science jobs, suggesting it may be somewhat harder to find remote jobs as a data analyst than other data science titles. However, remote job postings can also behave independently of total job postings, so remote jobs may at times become easier to find even while total jobs are dropping.


### 5. How do different schedule types make up total job postings?

Placing job counts as well as job count percentages by schedule types in a PivotTable, I could create a pie chart visualizing the proportions by which the different schedule types make up all job postings. I connected this table to the existing job title slicer.

### Analysis

* When including all data science jobs, full-time jobs are the majority with 27278 postings making up 77.01% of all postings, followed by contractor jobs at 14.24% and part-time at 5.9% with the rest being a slim minority.

#### What About Data Analysts?

* Full-time jobs are still the majority but make up a slightly smaller percentage with 7,647 being 73.68% of data analyst job postings. Contractor and part-time postings make up slightly higher percentages at 16.79% and 6.51% respectively.

## Building the Dashboard

For my dashboard, I thought the most helpful visualizations would include the earlier charts displaying top skills, yearly versus hourly salaries, and monthly job counts as well as job count comparisons by titles, top companies, and job counts by countries. In addition, while my previous analysis worksheets could only filter by job titles, I wanted my dashboard to allow for even more specific job filtering including by schedule types, remote jobs, and countries.

One of the limitations of Excel when it comes to data visualization is in the one-way relationships allowed to be defined in the data models controlling data flow. In order to relate and filter content upstream (job titles, salaries, etc) by variables downstream (skills and schedules), a two-way data relationship can be simulated using measures and a combination of DAX functions.

This measure is a job count that allows me to filter not only by job titles but also by schedule type:
```
=CALCULATE(
    DISTINCTCOUNT(data_jobs_skills[job_id]),
    TREATAS(
        VALUES(data_jobs_schedules[job_id]),
        data_jobs_skills[job_id]
    ))
```
I use this measure to calculate median salary and also filter by schedule type:
```
=IF(
    ISFILTERED(data_jobs_schedules[schedule_type]),
    CALCULATE(
        MEDIAN(data_jobs[salary_year_avg]),
        TREATAS(
            VALUES(data_jobs_schedules[job_id]),
            data_jobs[job_id]
        )
    ),
    MEDIAN(data_jobs[salary_year_avg])
)
```
Then I use these measures in PivotTables and PivotCharts to display the relationships I am looking for. The map chart, like the earlier histograms, could not be used directly from a PivotTable so I used dynamic arrays via the FILTER function to copy the data while leaving out blank cells to create the map chart:

```
=FILTER(K6:L112, K6:K112<>"")
```
Finally, I moved all the charts into a new worksheet and formatted the charts and worksheet to look visually clean and cohesive. As a last step, I protected all worksheets so nothing could be moved around or broken but still allow the slicers and filters to be used. I would also normally hide the Dashboard Analysis and Data Validation worksheets but I left them unhidden for the purposes of easier examining the work in this project.

## Final Analysis and Insights

[Final Dashboard]

* Data analysts are in highest demand as seen by having the highest count of job postings out of all data science titles followed by data scientist, data engineer, and their senior roles. Cloud engineers have the fewest postings.
* Data analyst jobs have the lowest median hourly pay and the second lowest median yearly salary. However, salary distribution seen in earlier histograms show some opportunities for much higher pay than the median.
* For all titles, yearly salaries tend to pay more than hourly salaries over the course of a year, sometimes by very large differences.
* It can become harder to find job opportunities as job postings become less frequent in the last months of the year.
* For someone pursuing a data analyst role, the most helpful skills to know appear to be SQL, Excel, Python, and Tableau as seen by their being in highest demand.
* The companies offering the most data science jobs are Robert Half, Insight Global, and Upwork. For data science jobs, Upwork is replaced by Get It Recruit.
* Most of the job postings in this dataset are located in the United States.

## Conclusion

By taking advantage of Power Query, Power Pivot, PivotTables, and a variety of functions and other Excel tools, a dataset of job postings was able to become a clear, cohesive, and visually pleasing summary of what I believed to be the most important relationships for anyone looking to decide the most important skills to learn, what kinds of salaries to look for, and the best times to find job opportunities depending on the data science role most important to the individual. And for anyone who doesn't yet know which data science role to focus on, the comparisons of salaries and job demand may also provide helpful insights towards making that decision.

Although I primarily focused on data analyst roles out of my personal interest whenever I decided to narrow in from the entirety of data science job postings, the dashboard and analysis worksheets in the `Job_Market_Analysis_Dashboard.xlsx` Excel workbook allow anyone to filter down to any job titles of interest. I hope this project can prove helpful to anyone else looking to learn more about the opportunities within data science jobs or about the functionality of Excel.
# Excel-Website-Performance-Analysis

## Project Overview
- Analyzed real GRAMMY.com and RecordingAcademy.com performance data using A/B testing to identify differences in audience behavior and engagement and develop data-driven recommendations for the Recording Academy’s digital strategy.

## About the dataset
- This dataset contains daily website traffic and user engagement data for two websites: the Grammys and The Recording Academy (TRA). It tracks important performance metrics, including unique visitors, pageviews, browsing sessions, bounced sessions, average session duration, and mobile visitors. Additionally, the dataset includes information about major events, such as the Grammy Awards and nominee announcements, as well as the type of content featured on each website. The purpose of the dataset is to analyze and compare website performance, identify visitor behavior patterns, and understand how major events and website content influence traffic and user engagement.
- The dataset contains two separate sections:
- Grammys: Tracks website traffic and visitor engagement specifically on the Grammys website.
- TRA (The Recording Academy): Tracks website traffic and visitor engagement on The Recording Academy website.

## Data Sample
- Formulas:
- This formula uses the IF function in Excel to split the Grammys data when Grammy.com began to only host Grammys content instead of both Recording Academy and Grammys content.
``` excel
=IF([@date] >= DATE(2022,2,1),"Grammys","Grammys + TRA" )
```
- This formula uses the XLOOKUP function in Excel to add the number of mobile visitors for the Grammys website after the website split in 2022.
``` excel
=XLOOKUP([@date],mobile_visits[date],mobile_visits[mobile_visitors],0)
```
- Grammys and Grammys + TRA table sample.
<img width="1000" height="370" alt="image" src="https://github.com/user-attachments/assets/3fdbeb81-e2c1-48ae-9bca-d314e9970d2c" />

- The Recording Academy table sample.
<img width="1000" height="370" alt="image" src="https://github.com/user-attachments/assets/5c207824-1f49-42bc-a135-9e081f4c807b" />

  




## Questions & Insights
Q1: Compare the average daily website visitors on days when an award ceremony was held to those when no awards ceremonies were held. 
- Based on the data, the Grammy Awards ceremony averages around 1,357,201 more visitors on the website than the average of regular day website visitors. This data reveals that even with the Grammy Awards appearing on the dataset 13 times out of 2,343 rows, their is still a very large difference in average website visitors. This shows how popular the ceremony truly is.
<img width="900" height="150" alt="image" src="https://github.com/user-attachments/assets/4cf9361f-8bd8-49e9-9ba8-126072fa41b9" />

Q2: What is the pages_per_session on the Grammys + TRA data?
- Formulas:
- This formula uses the SUMIF Excel function to calculate the total page views for the Grammys + TRA data.
``` excel
=SUMIFS(Grammys[pageviews],Grammys[website_content],"Grammys + TRA")
```
- This formula uses the SUMIF Excel function to calculate the total sessions for the Grammys + TRA data.
``` excel
=SUMIFS(Grammys[sessions],Grammys[website_content],"Grammys + TRA")
```
- Results:
<img width="466" height="190" alt="image" src="https://github.com/user-attachments/assets/a0c40160-a3d9-4365-bf18-e75cef08fd88" />

Q3: What is the website bounce rate for the Grammys website?

- Formulas:
``` excel
=SUMIFS(Grammys[bounced_sessions],Grammys[website_content], "Grammys")
```
``` excel
=SUMIFS(Grammys[sessions],Grammys[website_content],"Grammys")
```
- Results:
<img width="466" height="190" alt="image" src="https://github.com/user-attachments/assets/0d01468c-3c22-4646-82b2-71c62b82b0b2" />

Q4: What is the average time on website for the Grammys website compared to the Grammys + TRA website?

- Formulas:
``` excel
=AVERAGEIFS(Grammys[avg_session_duration_secs],Grammys[website_content],"Grammys")
```
``` excel
=AVERAGEIFS(Grammys[avg_session_duration_secs],Grammys[website_content],"Grammys + TRA")
```
- Results:
<img width="466" height="190" alt="image" src="https://github.com/user-attachments/assets/2fcb57e7-ace6-4feb-80a1-6835b3f49123" />

Q5: What percentage of users across all dates in the dataset use a mobile device to visit the website?
- Formulas:
``` excel
=SUMIFS(Grammys[mobile_visitors],Grammys[website_content],"Grammys")
```
``` excel
=SUMIFS(Grammys[visitors],Grammys[website_content],"Grammys")
```
- Results:
<img width="466" height="190" alt="image" src="https://github.com/user-attachments/assets/35d12c1e-90c0-48a8-bf86-f734c0ccbfc1" />

### Key findings & Recommendation
- Key Findings: The data overall suggests that the Recording Academy website is generating stronger user engagement than the Grammys site. The Recording Academy shows the highest average time on website with 128.5 seconds. The website also has the highest pages per session at 2.78 pages and the lowest website bounce rate at 34.7%. This suggests that users are exploring more content when visiting the Recording Academy website independently.
- Recommendation: The Recording Academy should keep the sites separate but create a stronger connection between both websites. Since The Recording Academy website has higher user engagement they could use that engagement to improve cross navigation between both websites, and share links with relevant content for both websites. This approach would preserve the benefits of having two specialized sites while making it easier for users to move between them when appropriate.







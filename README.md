# Interactive-sales-report-with-a-pivot-table

Sales Analysis – An interactive Excel report using a pivot table.

Notes: This is a project created for my portfolio, designed to help me practice and organize the knowledge I have acquired and to demonstrate the skills I have developed.

Why this project and data source: This is an analysis that is often performed in a real-world work environment, so it is a good way to test my skills in this context. The data comes from a live event on the YouTube channel hosted by Mr. Michał Kowalczyk, “Michał Kowalczyk – Excellent Work.” I will include the link below. However, if the link stops working in the future, it is because Mr. Kowalczyk sometimes removes some of these events after a while, once interested viewers have had a chance to practice the topics he selected.
https://www.youtube.com/watch?v=QWMEjOvRCnM

In addition, for this project, I also wanted to practice working with the English version of Excel, since I usually work with the Polish version. Therefore, since I realize that I am less familiar with the English equivalents of the Polish function names in Excel, I will complete this project in English. To do this, I will first modify the source file so that the headers and some of the values are in English.

 
<img width="1868" height="1000" alt="plik źródłowy" src="https://github.com/user-attachments/assets/db2f55f6-fad5-4dd1-8e62-b45314477a4c" />


Project objective – analyze sales over time, identify the top 5 stores by sales, and perform a Pareto analysis of sales value by warehouse.

1.	The data preparation process.

I copy the data from the source file into a new Excel file. There are 7,992 rows and 16 columns with the following names: numer zamówienia, kanał sprzedaży, magazyn, data zamówienia, data wysyłki, data dostawy, zespół sprzedażowy, region, ID klienta, ID sklepu, nazwa produktu, sztuki, zniżka, cena sztuki, koszt sztuki, sprzedaż.

Then I translate the Polish names into English and enter them.


<img width="945" height="496" alt="image" src="https://github.com/user-attachments/assets/6581ed9c-5815-49f2-8bf0-078f5b43cc5c" />



Next, I change the formatting of the columns containing dates, because I want everything to be in English (specifically, U.S. English).


<img width="945" height="509" alt="image" src="https://github.com/user-attachments/assets/86bf5ac0-4149-4ffe-97d5-4ecf865ebb3a" />


This means I am changing the regional settings from Polish to English.


<img width="945" height="507" alt="image" src="https://github.com/user-attachments/assets/6bc0f175-8133-43f2-8c65-6c0905f15cb8" />


Right away, to avoid calculation errors caused by different decimal separators, I replace the comma with a period.


<img width="945" height="505" alt="image" src="https://github.com/user-attachments/assets/b466990f-7d98-4c4d-881d-0a49d76cd8d4" />


The data has already been transformed and looks like a real dataset from an American store.


<img width="945" height="502" alt="image" src="https://github.com/user-attachments/assets/f38c8f81-3548-4450-be68-1ad4b0768d1e" />


Looking at the guidelines, I see that I have to present sales figures by warehouse. Therefore, I can shorten the warehouse name so it does not take up as much space on the chart. The prefix “WARE-” is the same for all of them, so I can remove it.


<img width="945" height="509" alt="image" src="https://github.com/user-attachments/assets/3b3af2db-0932-4f36-90ba-f711361d3e7d" />


I am adding a new column.


<img width="945" height="508" alt="image" src="https://github.com/user-attachments/assets/1335525f-0388-4331-b63d-7ca91e08cde8" />


Next, I use the =TEXTAFTER() function to extract the text that follows the delimiter. All that remains is to create a table from the range, which will allow me to refresh the range so that the pivot table is always up to date.


<img width="945" height="508" alt="image" src="https://github.com/user-attachments/assets/aaa83cf0-fd23-482e-bc3f-867536abf283" />


This is what the finished table looks like.


<img width="945" height="505" alt="image" src="https://github.com/user-attachments/assets/c1430755-15e6-430a-b658-e9c7069d6b3e" />


2.	The process of analyzing sales data.

I am creating a new worksheet in which I will include a pivot table showing sales data over time.



<img width="945" height="513" alt="image" src="https://github.com/user-attachments/assets/8b8042b0-aea4-40ac-866f-3ae72dd0cd20" />


<img width="945" height="511" alt="image" src="https://github.com/user-attachments/assets/620d05ec-59a2-44b7-b605-ba1a14148a34" />


I select “Sales" for the value fields and “Order Date” for the row fields.


<img width="945" height="509" alt="image" src="https://github.com/user-attachments/assets/6535345e-9bf6-432b-8d94-6f2389707a5e" />


I am inserting a line chart that I will later include in the report. I am setting it to show data by quarter.


<img width="945" height="515" alt="image" src="https://github.com/user-attachments/assets/42a52f09-202f-4dbb-beaa-c280da56cc85" />


Next, I will create a new worksheet listing the top 5 stores by sales.


<img width="945" height="532" alt="image" src="https://github.com/user-attachments/assets/d505af6e-5aee-4f6d-8a4c-224d757b29e9" />



I paste the pivot table from the previous worksheet and adjust it to find the answer to the next question.


<img width="945" height="511" alt="image" src="https://github.com/user-attachments/assets/1d2bffd1-0d21-40a8-8eca-12e9d6fa7122" />


Sales remain in the “Values” column, and I drag the Store ID to the rows.


<img width="945" height="516" alt="image" src="https://github.com/user-attachments/assets/4b21d4fb-8bf0-499b-8c1f-048783a31965" />


Then I select the filter option, which you can find here under “Top 10.”


<img width="945" height="508" alt="image" src="https://github.com/user-attachments/assets/6de29aef-3602-4463-b63d-aa41b32058f6" />


I set the number to the top 5 stores by sales.


<img width="945" height="508" alt="image" src="https://github.com/user-attachments/assets/c96061bb-473e-4732-a1cb-df323b763e53" />


I am inserting a column chart.


<img width="945" height="507" alt="image" src="https://github.com/user-attachments/assets/80b94282-8820-49dc-8328-146640dd630f" />


After that, I sort the values from largest to smallest to make the chart easier to read.


<img width="945" height="511" alt="image" src="https://github.com/user-attachments/assets/5f2011e8-1da5-4206-9590-97e6b464b354" />


The chart is now ready.

The next step is to perform a Pareto analysis of sales figures by warehouse.



<img width="945" height="507" alt="image" src="https://github.com/user-attachments/assets/9b464686-6a6e-4171-b4ac-d55fc3fb980c" />



First, I create a new worksheet and copy the pivot table.



<img width="945" height="506" alt="image" src="https://github.com/user-attachments/assets/c4e75f13-26ba-433b-9795-9d0fb075a15e" />


I leave “Sales” in the ‘Value’ field and drag the warehouse number into the “Rows” field.

To create a Pareto analysis, I need to know the percentage share of sales by warehouse. Therefore, I drag the sales data into the “Value” field once again.



<img width="945" height="511" alt="image" src="https://github.com/user-attachments/assets/24428607-0d8d-41e0-8515-a91bf6812764" />


I do not need to calculate anything, because the options allow me to calculate the running percentage based on my selection—in this case, the warehouse.


<img width="945" height="509" alt="image" src="https://github.com/user-attachments/assets/5f4b5634-e815-450c-80f0-15faa3d4ffeb" />


<img width="945" height="514" alt="image" src="https://github.com/user-attachments/assets/d5ff6e10-2e3a-4be8-a303-a2c8e364ccc3" />


I also sort the sales data from largest to smallest.


<img width="945" height="515" alt="image" src="https://github.com/user-attachments/assets/09d15de8-0b88-4193-8f7d-8e89092d89a5" />


I change the label names in the table and insert the appropriate chart.


<img width="945" height="508" alt="image" src="https://github.com/user-attachments/assets/8c2ecebf-c413-44c4-8396-b3bcf38f4943" />


I format the chart to make it easier to read.


<img width="945" height="511" alt="image" src="https://github.com/user-attachments/assets/e5f6f870-faf0-4231-966d-6060de9f7220" />


The Pareto chart is now ready.


<img width="945" height="505" alt="image" src="https://github.com/user-attachments/assets/19c82736-98fc-4778-b268-c704725ca265" />



According to the Pareto theory, 20% of warehouses should generate 80% of sales. As shown in the chart above, this is not an ideal situation according to the Pareto principle. The first warehouse accounts for about 30% of sales, while the two warehouses with the highest sales account for about 50% of products sold.


3.	The process of creating an interactive sales report.

I am looking at the source data to see exactly what this company sells and tailor the dashboard to the audience. These are mainly furniture, electronics, and home accessories.
I am searching for a suitable illustration to use as the background for my report.


<img width="945" height="504" alt="image" src="https://github.com/user-attachments/assets/d10e1f0f-71e8-4296-975f-2e39863b33c0" />


I select an image and adjust its size.


<img width="875" height="474" alt="image" src="https://github.com/user-attachments/assets/bef2ba96-f51f-4639-8510-fd4aae2e4b3d" />


I create fields in which I will place particular charts.


<img width="945" height="516" alt="image" src="https://github.com/user-attachments/assets/aa7db273-81c4-48f1-9f4a-3282662ded45" />


I set the fields and align them.


<img width="945" height="507" alt="image" src="https://github.com/user-attachments/assets/2fe038e4-5ab8-425b-bdf1-68474586769f" />


I add icons and label each field so that it matches the chart.


<img width="945" height="510" alt="image" src="https://github.com/user-attachments/assets/1e35fe52-d708-4a8f-b3fe-1e83f006840c" />


Using the Ctrl + X and Ctrl + V shortcuts, I paste the charts into the finished template.


<img width="945" height="513" alt="image" src="https://github.com/user-attachments/assets/cbd3e972-98ff-4702-825f-6e60674167a4" />


Next, I format each chart so that it is easy to read.


<img width="945" height="509" alt="image" src="https://github.com/user-attachments/assets/201fb0aa-e5eb-4915-bfdb-85b2686ded3f" />


To avoid wasting time formatting each chart one by one, I save this finished chart as a template.


<img width="945" height="511" alt="image" src="https://github.com/user-attachments/assets/db306f2c-5859-482b-9afb-1cb84da6e7e8" />


Then I use this template to format the remaining charts.


<img width="945" height="508" alt="image" src="https://github.com/user-attachments/assets/62a381f0-0971-4afc-88b3-751700444f12" />


After formatting, the charts look like this.


<img width="945" height="504" alt="image" src="https://github.com/user-attachments/assets/b9086776-8b0b-4a44-b95c-3140e13679c9" />


I set up a filter in the form of a slicer that will filter the results by sales channel.


<img width="945" height="510" alt="image" src="https://github.com/user-attachments/assets/93e5955a-ade9-41b4-8503-b605272888ba" />


The next step is to link the slicer to all the charts.


<img width="945" height="502" alt="image" src="https://github.com/user-attachments/assets/da7ab570-5c6e-4fce-92f9-1233ed5c2d42" />


I set the slicer to the appropriate format so that it matches the rest of the report, and the dashboard is ready.


<img width="945" height="511" alt="image" src="https://github.com/user-attachments/assets/ae05c1d8-22aa-4615-aba1-c8ea0d9a914c" />


The interactive sales report in Excel is ready. You can view sales results by sales channel, and if new data is added in the future, simply refresh the report to update the values.


<img width="945" height="368" alt="image" src="https://github.com/user-attachments/assets/0a5b36e7-972d-4dd8-ad69-a2ea32be61e9" />


Conclusions:
1.    Sales are consistent over time; there is no seasonality. There are no significant drops in sales based on sales channel.
2.	The 5 stores with the best sales performance are stores numbered: 284, 26, 238, 166, and 33. Stores numbered 284 and 26 consistently rank high in terms of sales across two other sales channel categories. The remaining stores do not have such stable sales across different sales channels.
3.	There is significant potential to achieve the Pareto principle in the future. Warehouse NMK1003 generates the highest sales, accounting for approximately 30% of total sales. The top two warehouses - NMK1003 and PUJ1005 - are responsible for approximately 50% of the company’s total sales. In the future, the company could invest more in developing the top four warehouses and discontinue operations at the bottom two (NBV1002 and MKL1006), which generate the least sales.










































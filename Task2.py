import pandas as pd
from xlsxwriter import Workbook
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def analyze_top_selling_categories():
    # Read the CSV file into a DataFrame
    df = pd.read_csv('review_dataset.csv')

    # Count the occurrences of each product category
    category_counts = df['category'].value_counts()

    # Display the top-selling product categories analysis
    print("Top Selling Product Categories:")
    print(category_counts.head())
    result = pd.DataFrame(category_counts.head())
    return result

def analyze_top_consumer_cities():
    # Read the CSV file into a DataFrame
    df = pd.read_csv('orders_2016-2020_Dataset.csv')

    # Filter orders with billing country as India
    df_india = df[df['Billing Country'] == 'IND']

     # Convert city values to lowercase
    df_india['Billing City'] = df_india['Billing City'].str.lower()

    # Count the occurrences of each city
    city_counts = df_india['Billing City'].value_counts()

    # Display the top consumer cities analysis
    print("Top Consumer Cities in India:")
    print(city_counts.head())
    result = pd.DataFrame(city_counts.head())
    return result

def analyze_top_consumer_states():
    # Read the CSV file into a DataFrame
    df = pd.read_csv('orders_2016-2020_Dataset.csv')

    # Filter orders with billing country as India
    df_india = df[df['Billing Country'] == 'IND']

    # Count the occurrences of each state
    state_counts = df_india['Billing State'].value_counts()

    # Display the top consumer states analysis
    print("Top Consumer States in India:")
    print(state_counts.head())
    result = pd.DataFrame(state_counts.head())
    return result

def analyze_payment_methods():
    # Read the CSV file into a DataFrame
    df = pd.read_csv('orders_2016-2020_Dataset.csv')

    # Extract the 'Payment Method' column
    payment_methods = df['Payment Method']

    # Count the occurrences of each payment method
    payment_method_counts = payment_methods.value_counts()

    # Display the payment method analysis
    print("Payment Method Analysis:")
    print(payment_method_counts)

    # Generate report
    pdf_file = 'Payment_Method_Analysis_report.pdf'
    c = canvas.Canvas(pdf_file, pagesize=letter)

    # Set up report content
    report_title = "Payment Method Analysis"
    total_reviews_text = f"{payment_method_counts}"
    # Draw report content on the canvas
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, 750, report_title)
    c.setFont("Helvetica", 12)
    t=c.beginText()
    t.setFont('Helvetica-Bold', 10)
    t.setCharSpace(3)
    t.setTextOrigin(50, 700)
    t.textLines(total_reviews_text)
    c.drawText(t)
    # Save the canvas as PDF
    c.save()
    # Generate report DataFrame
    # Display report
    print(f"Payment Method Analysis report generated as '{pdf_file}'")
    report_data={
         "Payment Method Analysis" : [payment_method_counts]
     }
    df_report = pd.DataFrame(report_data)
    df_report.to_excel("Payment_Method_Analysis_report.xlsx", index=False, na_rep=' ')
    # Display report
    print("Payment Method Analysis report generated as 'Payment_Method_Analysis_report.xlsx'")
    return  df_report

def parse_rating(series):
    # Filter out null values
    series = series.dropna()
    # Remove " star rating" from each element in the series
    rating_text = series.str.replace(" star rating", "")
    
    # Parse the rating values as floats
    ratings = rating_text.astype(float)
    
    return ratings



def analyze_reviews():
    # Read the CSV file into a DataFrame
    df = pd.read_csv('review_dataset.csv')
    
    # Perform analysis
    total_reviews = len(df.dropna())
    total_stars =parse_rating(df['stars']).sum()
    average_rating = total_stars / total_reviews
    
    print("Analysis of Reviews:")
    print("--------------------")
    print("Total Reviews:", total_reviews)
    print("Average Rating:", average_rating)
    print("--------------------")
    print("Individual Reviews:")
    print(df.dropna())
    print("About data")
    print(df.describe())
    print("Most reviewed item:")
    print(df["product_name"].value_counts().idxmax()+" \nCount="+ str(df['product_name'].value_counts().loc[df["product_name"].value_counts().idxmax()]))
    print("Least reviewed item:")
    print(df["product_name"].value_counts().idxmin()+" \nCount="+ str(df['product_name'].value_counts().loc[df["product_name"].value_counts().idxmin()]))
    print("Most bought category:")
    print(df['category'].value_counts().idxmax()+" \nCount="+ str(df['category'].value_counts().loc[df['category'].value_counts().idxmax()]))
    print("Least bought category:")
    print(df['category'].value_counts().idxmin()+" \nCount="+ str(df['category'].value_counts().loc[df['category'].value_counts().idxmin()]))
    print("--------------------")
    # Add any additional analysis or visualization logic here
    # Generate visualization
    plt.figure(figsize=(8, 6))
    plt.hist(parse_rating(df['stars']), bins=5, edgecolor='black')
    plt.xlabel('Rating')
    plt.ylabel('Frequency')
    plt.title('Distribution of Ratings')
    plt.savefig('histogram.png')

    # Generate report
    pdf_file = 'review_analysis_report.pdf'
    c = canvas.Canvas(pdf_file, pagesize=letter)

    # Set up report content
    report_title = "Review Analysis Report"
    total_reviews_text = f"Total Reviews: {total_reviews}"
    average_rating_text = f"Average Rating: {average_rating}"
    aboutdata = f"About data: {df.describe()}"
    mostreviewed = f'Most reviewed item: {df["product_name"].value_counts().idxmax()}'
    mostreviewedcount = f'Most reviewed count: {df["product_name"].value_counts().loc[df["product_name"].value_counts().idxmax()]}'
    leastreviewed = f'Least reviewed item: {df["product_name"].value_counts().idxmin()}'
    leastreviewedcount = f'Least reviewed count: {df["product_name"].value_counts().loc[df["product_name"].value_counts().idxmin()]}'
    mostbought = f'Most bought item: {df["category"].value_counts().idxmax()}'
    mostboughtcount = f'Most bought item count: {df["category"].value_counts().loc[df["category"].value_counts().idxmax()]}'
    leastbought = f'Least bought item: {df["category"].value_counts().idxmin()}'
    leastboughtcount = f'Least bought item count: {df["category"].value_counts().loc[df["category"].value_counts().idxmin()]}'
    # Draw report content on the canvas
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, 750, report_title)
    c.setFont("Helvetica", 12)
    c.drawString(50, 700, total_reviews_text)
    c.drawString(50, 680, average_rating_text)
    c.drawString(50, 660, mostreviewed)
    c.drawString(50, 640, mostreviewedcount)
    c.drawString(50, 620, leastreviewed)
    c.drawString(50, 600, leastreviewedcount)
    c.drawString(50, 580, mostbought)
    c.drawString(50, 560, mostboughtcount)
    c.drawString(50, 540, leastbought)
    c.drawString(50, 520, leastboughtcount)
    # Save the canvas as PDF
    c.save()

    # Display report
    print(f"Review analysis report generated as '{pdf_file}'")

     # Generate report DataFrame
    report_data = {
        'Total Reviews': [total_reviews],
        'Average Rating': [average_rating],
        "Most reviewed item:" : [df["product_name"].value_counts().idxmax()],
        'Most reviewed item count' : [df['product_name'].value_counts().loc[df["product_name"].value_counts().idxmax()]],
        "Least reviewed item:" : [df["product_name"].value_counts().idxmin()],
        "Least reviewed item count:" : [df['product_name'].value_counts().loc[df["product_name"].value_counts().idxmin()]],
        "Most bought category:" : [df["category"].value_counts().idxmax()],
        "Most bought category count:" : [df['category'].value_counts().loc[df["category"].value_counts().idxmax()]],
        "Least bought category:" : [df["category"].value_counts().idxmin()],
        "Least bought category count:" : [df['category'].value_counts().loc[df["category"].value_counts().idxmin()]]
    }
    report_df = pd.DataFrame(report_data)

    # Save report as Excel file
    report_df.to_excel('review_analysis_report.xlsx', index=False)

    # Display report
    print("Review analysis report generated as 'review_analysis_report.xlsx'")
    return report_df

def analyze_reviews_for_categories():
    # Read the CSV file into a DataFrame
    df = pd.read_csv('review_dataset.csv')

    # Parse the rating from the 'stars' column
    df['rating'] = parse_rating(df['stars'])

    # Group the reviews by product category and calculate the average rating
    category_ratings = df.groupby('category')['rating'].mean()

    # Display the analysis of reviews for all product categories
    print("Analysis of Reviews for All Product Categories:")
    print(category_ratings)
    result = pd.DataFrame(category_ratings)
    return result

def analyze_orders_per_month_per_year():
    # Read the CSV file into a DataFrame
    df = pd.read_csv('orders_2016-2020_Dataset.csv')

    # Convert the 'Order Date and Time Stamp' column to datetime type
    df['Order Date and Time Stamp'] = pd.to_datetime(df['Order Date and Time Stamp'])

    # Extract the year and month from the 'Order Date and Time Stamp' column
    df['Year'] = df['Order Date and Time Stamp'].dt.year
    df['Month'] = df['Order Date and Time Stamp'].dt.month

    # Group the orders by year and month and count the number of orders
    orders_per_month_per_year = df.groupby(['Year', 'Month'])['Order #'].count()

    # Display the analysis of the number of orders per month per year
    print("Analysis of Number of Orders Per Month Per Year:")
    print(orders_per_month_per_year)
    result = pd.DataFrame(orders_per_month_per_year)
    return result

def analyze_reviews_for_orders_per_month_per_year():
    # Load the review dataset
    reviews_df = pd.read_csv('orders_2016-2020_Dataset.csv')

    # Convert the 'Order Date and Time Stamp' column to datetime
    reviews_df['Order Date and Time Stamp'] = pd.to_datetime(reviews_df['Order Date and Time Stamp'])

    # Extract the year and month from the 'Order Date and Time Stamp' column
    reviews_df['Year'] = reviews_df['Order Date and Time Stamp'].dt.year
    reviews_df['Month'] = reviews_df['Order Date and Time Stamp'].dt.month

    # Group the reviews by year, month, and count the number of reviews
    reviews_per_month_per_year = reviews_df.groupby(['Year', 'Month'])['Order #'].count().reset_index()

    # Plot the reviews per month per year
    plt.figure(figsize=(10, 6))
    plt.plot(reviews_per_month_per_year['Month'], reviews_per_month_per_year['Order #'], marker='o')
    plt.xlabel('Month')
    plt.ylabel('Number of Reviews')
    plt.title('Reviews for Number of Orders Per Month Per Year')
    plt.xticks(range(1, 13))
    plt.grid(True)
    # Create a DataFrame from the plot data
    plot_data = pd.DataFrame({
        'Month': reviews_per_month_per_year['Month'],
        'Number of Reviews': reviews_per_month_per_year['Order #']
    })
    plt.show()
    return plot_data


def menu():
    print("Input Value to Generate Graph Chart:")
    print("Enter 1 to see the analysis of Reviews given by Customers")
    print("Enter 2 to see the analysis of different payment methods used by Customers")
    print("Enter 3 to see the analysis of Top Consumer States of India")
    print("Enter 4 to see the analysis of Top Consumer Cities of India")
    print("Enter 5 to see the analysis of Top Selling Product Categories")
    print("Enter 6 to see the analysis of Reviews for All Product Categories")
    print("Enter 7 to see the analysis of Number of Orders Per Month Per Year")
    print("Enter 8 to see the analysis of Reviews for Number of Orders Per Month Per Year")
    print("Enter 9 to see the analysis of Number of Orders Across Parts of a Day")
    print("Enter 10 to see the Full Report")
    print("Enter the number to see the analysis of your choice: ")

def analyze_orders_across_parts_of_day():
    # Read the CSV file into a DataFrame
    df = pd.read_csv('orders_2016-2020_Dataset.csv')

    # Convert the 'Order Date and Time Stamp' column to datetime type
    df['Order Date and Time Stamp'] = pd.to_datetime(df['Order Date and Time Stamp'])

    # Extract the hour from the 'Order Date and Time Stamp' column
    df['Hour'] = df['Order Date and Time Stamp'].dt.hour

    # Assign parts of a day based on the hour
    df['Part of Day'] = pd.cut(df['Hour'], bins=[0, 6, 12, 18, 24], labels=['Night', 'Morning', 'Afternoon', 'Evening'])

    # Group the orders by part of day and count the number of orders
    orders_per_part_of_day = df.groupby('Part of Day')['Order #'].count()

    # Display the analysis of the number of orders across parts of a day
    print("Analysis of Number of Orders Across Parts of a Day:")
    print(orders_per_part_of_day)
    result = pd.DataFrame(orders_per_part_of_day)
    return result

def generate_full_report():
    # Read the required CSV files into DataFrames
    reviews_df = pd.read_csv('review_dataset.csv')
    orders_df = pd.read_csv('orders_2016-2020_Dataset.csv')

    # Perform the analysis for each option
    analysis_options = {
        'Reviews Analysis': analyze_reviews(),
        'Payment Methods Analysis': analyze_payment_methods(),
        'Top Consumer States Analysis': analyze_top_consumer_states(),
        'Top Consumer Cities Analysis': analyze_top_consumer_cities(),
        'Top Selling Product Categories Analysis': analyze_top_selling_categories(),
        'Reviews for All Product Categories Analysis': analyze_reviews_for_categories(),
        'Number of Orders Per Month Per Year Analysis': analyze_orders_per_month_per_year(),
        'Reviews for Number of Orders Per Month Per Year Analysis': analyze_reviews_for_orders_per_month_per_year(),
        'Number of Orders Across Parts of a Day Analysis': analyze_orders_across_parts_of_day()
    }

    # Create a new Excel writer object
    writer = pd.ExcelWriter('full_report.xlsx', engine='xlsxwriter')

    # Iterate over the analysis options and save each analysis result in a separate worksheet
    for analysis_name, analysis_result in analysis_options.items():
        if analysis_result is not None:
            worksheet_name = analysis_name[:31]
            analysis_result.to_excel(writer, sheet_name=worksheet_name, index=True)
    # Save and close the Excel writer
    writer.save()
    writer.close()

    print("Full report generated successfully as 'full_report.xlsx'.")


def main():
    menu()
    choice = int(input())
    
    if choice == 1:
        # Analysis of Reviews given by Customers
        print("Performing analysis of Reviews given by Customers...")
        # Add your code for this analysis here
        analyze_reviews()
        print("OutPut:Genrate analysis report in format PDF and Excel file.")
        
    elif choice == 2:
        # Analysis of different payment methods used by Customers
        print("Performing analysis of different payment methods used by Customers...")
        # Add your code for this analysis here
        analyze_payment_methods()
        
    elif choice == 3:
        # Analysis of Top Consumer States of India
        print("Performing analysis of Top Consumer States of India...")
        # Add your code for this analysis here
        analyze_top_consumer_states()
        
    elif choice == 4:
        # Analysis of Top Consumer Cities of India
        print("Performing analysis of Top Consumer Cities of India...")
        # Add your code for this analysis here
        analyze_top_consumer_cities()
        
    elif choice == 5:
        # Analysis of Top Selling Product Categories
        print("Performing analysis of Top Selling Product Categories...")
        # Add your code for this analysis here
        analyze_top_selling_categories()
        
    elif choice == 6:
        # Analysis of Reviews for All Product Categories
        print("Performing analysis of Reviews for All Product Categories...")
        # Add your code for this analysis here
        analyze_reviews_for_categories()
        
    elif choice == 7:
        # Analysis of Number of Orders Per Month Per Year
        print("Performing analysis of Number of Orders Per Month Per Year...")
        # Add your code for this analysis here
        analyze_orders_per_month_per_year()
        
    elif choice == 8:
        # Analysis of Reviews for Number of Orders Per Month Per Year
        print("Performing analysis of Reviews for Number of Orders Per Month Per Year...")
        # Add your code for this analysis here
        analyze_reviews_for_orders_per_month_per_year()
        
    elif choice == 9:
        # Analysis of Number of Orders Across Parts of a Day
        print("Performing analysis of Number of Orders Across Parts of a Day...")
        # Add your code for this analysis here
        analyze_orders_across_parts_of_day()
        
    elif choice == 10:
        # Full Report
        print("Generating the Full Report...")
        # Add your code for the full report here
        generate_full_report()
        
    else:
        print("Invalid choice. Please try again.")

if __name__ == '__main__':
    main()

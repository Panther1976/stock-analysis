<style>
    .page-content img {
        border: 2px solid #DDD;
    }
</style>
# **Stock Analysis**

## **Overview of Project**

The original stock that our client's parents were interested in investing in turned out to be a poor investment choice. So our client has asked to help evaluate a list of green stocks based on trade volume. The purpose of this project is to use the data collected to allow our client to make a more informed decision. 

### **Results**

The Purpose of this analysis is to evaluate potential investment opportunities.

### **Performance for 2017**

In 2017, The top 2 performers based on return were DQ at 199.4% and SEDG at 184.5%. 

![2017 Returns](https://user-images.githubusercontent.com/106631875/175757686-5d6f04e9-9102-4304-b2e2-9c21558b946d.png)

### **Performance for 2018**

![2018 Returns](https://user-images.githubusercontent.com/106631875/175757688-d7662970-1ac5-49cd-864a-2219f3fbcdb2.png)

In 2018, The top 2 performers based on return were ENPH at 81.9% and RUN at 84.0%. 

The top 2 performers in 2017 posted returns of -62.6% for DQ and -7.8% for SEDG. 


### **Summary**

-	What are the advantages or disadvantages of refactoring code?

The advantage of refactoring code is making it more efficient. By taking less step to do the same functions, time is saved. This module introduced the concept of DRY (Don't Repeat Yourself). So instead of copying and pasting the code for the DQ ticker for each of the other tickers, we re-wrote the code using a loop. See below.

![All_Tickers_Loop](https://user-images.githubusercontent.com/106631875/178597834-a9247de5-d019-4848-8906-f9f5557fa1db.png)

-	How do these pros and cons apply to refactoring the original VBA script?

In my first attempt, I re-used some of the coding I created during the module instead of starting with the refactored code given. The result was that I had nested loops that were slowing the speed of the macro. After starting over with the refactored code that was given, I was able to correct my mistakes. This allowed me to make the code more efficient which resulted much faster results. My original results took several seconds. After correcting my mistakes, the results take a fraction of a second. See below.

#### Time Elapsed to Run the 2017 Stock Macro:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/106631875/178599270-73aa0df1-f398-4a9f-a358-b65cafd44d5f.png)

#### Time Elapsed to Run the 2018 Stock Macro:

![VBA_Challenge_2018](https://user-images.githubusercontent.com/106631875/178599315-6ed575d1-d46a-4eb7-8200-4620ed634909.png)


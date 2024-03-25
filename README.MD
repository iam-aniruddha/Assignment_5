# Automation Tasks with Selenium and Openpyxl

## Problem Statements

### Task 1: Retrieve User Credentials

Retrieve all the Users and password related information and store it in Sheet named "User credentials" in a excel file.
User ID, User Name, Password

### Task 2: Login and Error Recording

Try login with all the Users provided on the website link https://www.saucedemo.com/v1/ and record any error messages
displayed for a specific user. write this information to another sheet named "Login" in the same excel file as above
in the following format.
User ID, Login Message

### Task 3: Retrieve Product Details

Login using the "standard_user" and retrieve all the product related information for every listed product.
Write this information into a separate sheet named "Product Details" in the same excel file.
Product ID, Product Name, Description, Price

### Task 4: Manual Order Details Entry

Manually create another sheet called "Order Details" in the same excel file and add Orders for
the "standard_user" and "problem_user". Decide on the Columns on your own and try to represent this data.
(review before you move to next one)

### Task 5: Placing Orders

Using the above "Order Details" sheet try the place all the orders on the website one by one. Update a new column in the
"Order Details" sheet called "Order Status" marking it as "Success/Failure".
Success criteria should be as follows.

1. The correct number of items should be added to the Cart.
2. The total amount at the end should be correct.
3. Success message should be displayed for every order placed.
4. No Items should be left over in the cart.

## Solution Approach

1. **Task 1 - Retrieve User Credentials**:

   - Utilize Selenium to extract user credentials from the website.
   - Write the extracted data to an Excel file using Openpyxl.
2. **Task 2 - Login and Error Recording**:

   - Use Selenium to automate the login process for each user.
   - Record any error messages displayed during the login attempt.
   - Write the login status (success or error message) to the Excel file.
3. **Task 3 - Retrieve Product Details**:

   - Log in with a specific user using Selenium.
   - Extract product details for all listed products on the website.
   - Write the product details to the Excel file.
4. **Task 4 - Manual Order Details Entry**:

   - Manually create and add orders for specific users in the Excel sheet.
   - Write order details such as order ID, user ID, product name, quantity, and total price.
5. **Task 5 - Placing Orders**:

   - Use Selenium to place orders on the website using the order details from the Excel sheet.
   - Update the order status based on success criteria (correct number of items added, total amount, success message, empty cart).

## Dependencies

- Python 3.x
- Selenium: `pip install selenium`
- Openpyxl: `pip install openpyxl`
- Chrome WebDriver: Download from [here](https://chromedriver.chromium.org/downloads) and ensure it's in your PATH.

## Usage

1. Clone this repository to your local machine.
2. Install the required dependencies listed above.
3. Update any file paths or website URLs in the code if necessary.
4. Run each Python script corresponding to the task you want to execute.
5. Check the Excel file generated after running each task for the desired output.

## Additional Notes

- Ensure an active internet connection while running tasks involving web automation.
- Verify that the scripts handle different scenarios effectively and provide the expected output.
- Adapt and modify the code as per specific requirements or website configurations.

![loginm](https://user-images.githubusercontent.com/76075516/122144833-99402c00-ce2a-11eb-97ec-363c7accb988.png)


>status: development⚠️

### It is a desktop application planned by me, where I perform the CRUD of the processes carried out in the system in addition to other features, such as generating a report, among others, which is used in a small business.

## Some fields in the main template are:

+ customer registration
+ product registration
+ automatic address lookup via zip code
+ consultation of registered products
+ product category, stock, stock control
+ sales module, possible to generate sales report
+ automated email sending  
In addition, there is a user with these fields:

+ name
+ email
+ cpf
+ active

## In addition to CRUD, I implement other features, such as:

![principal](https://user-images.githubusercontent.com/76075516/122144896-b07f1980-ce2a-11eb-96b1-706d14f80d58.png)


* zip query via API (pycepmail)
* Complete verification system to validate customer and user forms.
* Success message when creating, editing, registering among other features.
* Profile editable by adm user.
* User registration.
* Sales, being able to audit the sales made in order to export the reports generated in Excel.

![enviamail](https://user-images.githubusercontent.com/76075516/122145309-68acc200-ce2b-11eb-9cff-127ea314299e.png)

* Sending emails using outlook, with process automation and html formatting in the backend

![tsds](https://user-images.githubusercontent.com/76075516/122145181-22eff980-ce2b-11eb-917c-5c17b9286514.png)


## These features are under development:

- Further improve the sales screen to generate notes
- Change email server to gmail.

## Technologies used:

<table>
  <tr>
    <td> Python </td>
    <td> PQt5 </td>
    <td> MySql </td>
  </tr>
  <tr>
    <td> 3.7.0 * </td>
    <td>5.0</td>
    <td>8.0</td>
  </tr>
</table>

## How to run the application:

1) It can be run in a development environment, the libraries used in the project will be charged, but if the repository is cloned, all the libraries are already installed, and dependencies can be installed if necessary through the 'pip' package manager of python.

2) All dependencies are in the project body, including the DB
3) create a new MySql Schema
4) create the virtual environment with virtualenv if necessary
5) configure the DB connection string in the python script or as you prefer
6) After performing all the steps, just with the virtual environment activated, run the .py file


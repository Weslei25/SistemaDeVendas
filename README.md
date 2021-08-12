![Screenshot_2](https://user-images.githubusercontent.com/76075516/129206312-5b632f1a-3d3b-4300-9e8e-b1a60013e684.png)



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

![Screenshot_1](https://user-images.githubusercontent.com/76075516/129206059-9d66bfb5-be29-4e5f-b945-891ee2fdb7b2.png)



* zip query via API (pycepmail)
* Complete verification system to validate customer and user forms.
* Success message when creating, editing, registering among other features.
* Profile editable by adm user.
* User registration.
* Sales, being able to audit the sales made in order to export the reports generated in Excel.

![Screenshot_3](https://user-images.githubusercontent.com/76075516/129206471-5489ed0f-0f88-4211-adbe-f49ebfafb7c2.png)


* Sending emails using outlook, with process automation and html formatting in the backend

![Screenshot_4](https://user-images.githubusercontent.com/76075516/129206655-342883fd-59c1-4523-b055-23db18036bbb.png)


* New screen for company registration and consultation with the integration in the api "https://receitaws.com.br/api" ,
to facilitate the registration of companies.

![Screenshot_5](https://user-images.githubusercontent.com/76075516/129207043-6e7fd53c-fd1d-4ad3-8034-a5319eec0b2a.png)
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


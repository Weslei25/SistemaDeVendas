<h1>Think_System</h1>

>status: development⚠️

### It is a desktop application planned by me, where I perform the CRUD of the processes carried out in the system in addition to other features, such as generating a report, among others, which is used in a small business.

## Some fields in the main template are:

+ name
+ description
+ repetition num
+ sequence num
+ hard category
+ i know
+ user_id
+ image
  
In addition, there is a user with these fields:

+ name
+ email
+ cpf
+ active

## In addition to CRUD, I implement other features, such as:

* zip query via API (pycepmail)
* Complete verification system to validate customer and user forms.
* Success message when creating, editing, registering among other features.
* Profile editable by adm user.
* User registration.
* Sales, being able to audit the sales made in order to export the reports generated in Excel.
* Sending emails using outlook, with process automation and html formatting in the backend

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
    <td> 8. * </td>
    <td>7.4</td>
    <td>2.0</td>
  </tr>
</table>

## How to run the application:

1) It can be run in a development environment, the libraries used in the project will be charged, but if the repository is cloned, all the libraries are already installed, and dependencies can be installed if necessary through the 'pip' package manager of python.

2) All dependencies are in the project body, including the DB
3) create a new MySql Schema
4) create the virtual environment with virtualenv if necessary
5) configure the DB connection string in the python script or as you prefer
6) After performing all the steps, just with the virtual environment activated, run the .py file


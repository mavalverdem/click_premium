# ClickPremium API

- [ClickPremium API](#clickpremium-api)
    - [Auth](#auth)
        - [Register](#register)
            - [Register Request](#register-request)
            - [Register Response](#register-response)
        - [Login](#login)
            - [Login Request](#login-request)
            - [Login Response](#login-response)

## Auth

### Register

```js
POST {{host}}/auth/register
```

#### Register Request

```json
{
    "firstName":"Julio",
    "LastName":"Ayasta",
    "email": "julioayasta@gmail.com",
    "password":"secret"
}
```
#### Register Response

```js
200 OK
```

```json
{
    "id":"d89c2d9a-eb3e-4075-95ff-b920b55aa104",
    "firstName":"Julio",
    "LastName":"Ayasta",
    "email": "julioayasta@gmail.com",
    "token":"eyJhb..hbbQ"
}
```

## Login 

#### Login Request

```json
{
    "email": "julioayasta@gmail.com",
    "password":"secret"
}
```

#### Login Response

```js
200 OK
```

```json
{
    "id":"d89c2d9a-eb3e-4075-95ff-b920b55aa104",
    "firstName":"Julio",
    "LastName":"Ayasta",
    "email": "julioayasta@gmail.com",
    "token":"eyJhb..hbbQ"
}
```
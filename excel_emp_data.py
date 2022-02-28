from excel_classes import Employee,Address
import random


city_state = {"pune":"mh","panjim":"goa","jaipur":"rajastan","delhi":"delhi"}


def generate_name():
    name=""
    for i in range(random.randint(4,8)):
        character = chr(65+random.randint(0,25))
        name += character

    return name.title()


def generate_emp(num):
    list_emp = []
    for i in range(1,num+1):
        cty = random.choice(list(city_state.keys()))

        emp = Employee(eid=100+i,ename=generate_name(),eage=random.randint(25,45),esalary=random.randint(15000,90000),
        address = Address(pin=random.randint(100000,999999),city=cty,state=city_state.get(cty)))
        list_emp.append(emp)

    return list_emp

# print(generate_emp(100))












class common:
    def __str__(self):
        return f"{self.__dict__}"

    def __repr__(self):
        return str(self)
        

class Employee:
    def __init__(self,eid,ename,eage,esalary,address):
        self.Empid=eid
        self.Empname=ename
        self.Empage=eage
        self.Empsalary=esalary
        self.Empaddress=address
        
class Address:
    def __init__(self,pin,city,state):
        self.pincode=pin
        self.City=city
        self.state=state

# if __name__=="__main__":
#     pass
        













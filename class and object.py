#class Student :
#    name="aslam"
#s1=Student()
#print(s1.name)

#s2=Student()
#print(s2.name)
#class car:
#    color="blue"
 #   brand="tata"
#car1 =car()
#print(car1.color)
#print(car.brand)
class Student:
    collagename="rvit"
    name="any"
    def __init__(self,fullname,marks):
        self.name=fullname
        self.marks=marks
    def Welcome(self):
        print("welcome students",self.name)
    def get_marks(self):
        return self.marks
print(Student.collagename)
s1=Student("aslam",99)
print(s1.name,s1.marks)
s2=Student("gokul",100)
print(s2.name,s2.marks)
s1.Welcome()
print(s1.get_marks())

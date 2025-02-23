class Student:
    def __init__(self,name,marks):
        self.name=name
        self.marks=marks
    def avg(self):
        sum=0
        for val in self.marks:
            sum+=val
        print("hi",self.name,"your avg mark is:",sum/3)
s1=Student("aslam",[99,98,99])
s1.avg()
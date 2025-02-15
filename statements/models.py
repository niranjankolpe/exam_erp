from django.db import models

# Create your models here.
class University(models.Model):
    id = models.AutoField(primary_key=True)
    name = models.CharField(max_length=100, null=False, unique=True)

    def __str__(self):
        return str(self.id)

class College(models.Model):
    id = models.AutoField(primary_key=True)
    name = models.CharField(max_length=100, null=False, unique=True)

    def __str__(self):
        return str(self.id)

class Department(models.Model):
    id = models.AutoField(primary_key=True)
    name = models.CharField(max_length=100, null=False, unique=True)

    def __str__(self):
        return str(self.id)

class Subject(models.Model):
    id = models.AutoField(primary_key=True)
    name = models.CharField(max_length=100, null=False, unique=True)

    def __str__(self):
        return str(self.id)

class Teacher(models.Model):
    id = models.AutoField(primary_key=True)
    name = models.CharField(max_length=100, null=False, unique=True)
    college = models.ForeignKey(College, on_delete=models.CASCADE, null=False)

    def __str__(self):
        return str(self.id)
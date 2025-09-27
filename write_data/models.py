from django.db import models

class Client(models.Model):
    client_name = models.CharField(max_length=255, null=False)
    client_contact = models.CharField(max_length=255, null=True)
    report_quantity = models.IntegerField(default=0, null=False)
    client_template_quantity = models.IntegerField(default=0, null=False)

    class Meta:
        managed = True
        db_table = "Client"


class Project(models.Model):

    client = models.ForeignKey(
        Client,
        on_delete=models.CASCADE,
        related_name="projects",
        null=True
    )
    report_quantity = models.IntegerField(default=0, null=False)

    class Meta:
        managed = True
        db_table = "Project"


class Employee(models.Model):
    employee_name = models.CharField(max_length=255, null=False)
    report_quantity_per_employee = models.IntegerField(default=0, null=False)
    employee_role = models.CharField(max_length=120, null=True)

    class Meta:
        managed = True
        db_table = "Employee"


class Location(models.Model):
    location_name = models.CharField(max_length=255, null=False)
    location_template_quantity = models.IntegerField(default=0, null=False)

    class Meta:
        managed = True
        db_table = "Location"


class Template(models.Model):
    template_name = models.CharField(max_length=255, null=False)
    client = models.ForeignKey(
        Client,
        on_delete=models.CASCADE,
        related_name="templates"
    )
    location = models.ForeignKey(
        Location,
        on_delete=models.CASCADE,
        related_name="templates"
    )


    class Meta:
        managed = True
        db_table = "Template"


class WaterType(models.Model):
    water_type = models.CharField(max_length=255, null = False)
    template = models.ForeignKey(
        Template,
        on_delete=models.CASCADE,
        related_name="water_types"
    )

    class Meta:
        managed = True
        db_table = "WaterType"

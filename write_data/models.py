from django.db import models

class Project(models.Model):

    client_name = models.CharField(max_length=255)
    report_quantity = models.IntegerField()

    class Meta:

        managed = True
        db_table = "project_table"




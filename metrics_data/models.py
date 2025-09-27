from django.db import models

# Create your models here.

class Reporter(models.Model):


    reporter_name = models.CharField(max_length=100)
    report_quantity = models.IntegerField()


    class Meta:

        managed = True
        db_table = "reporters"

from rest_framework import serializers

from write_data.models import Employee


class MetricsDataSerializer(serializers.ModelSerializer):

    class Meta:

        model= Employee
        fields = ['id', 'employee_name']



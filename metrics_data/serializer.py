from rest_framework import serializers

from metrics_data.models import Reporter
from write_data.models import Employee


class MetricsDataSerializer(serializers.ModelSerializer):

    class Meta:

        model= Reporter
        fields = ['id', 'reporter_name']



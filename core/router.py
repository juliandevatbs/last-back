from django.db import connection, connections


class DatabaseRouter:

    # CONNECTIONS FROM DIFFERENTS DATABASES

    def conn_sampler(self, query, params=None):

        with connections['default'].cursor() as cursor:
            cursor.execute(query, params or [])
            return cursor.fetchall()

    def conn_project(self, query, params=None):

        with connections['project'].cursor() as cursor:
            cursor.execute(query, params or [])
            return cursor.fetchall()
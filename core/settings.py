from pathlib import Path
from dotenv import load_dotenv
import os

# Build paths inside the project like this: BASE_DIR / 'subdir'.
BASE_DIR = Path(__file__).resolve().parent.parent

load_dotenv() # Load environment variables from .env file

SECRET_KEY = os.getenv("SECRET_KEY")


DEBUG = True

ALLOWED_HOSTS = [
    "127.0.0.1",
    "localhost",
]

CORS_ALLOWED_ORIGINS = [
    "http://127.0.0.1:5173",
    "http://localhost:5173",

]

CORS_ALLOW_CREDENTIALS = True

CORS_ALLOWED_HEADERS = [
    'accept',
    'accept-encoding',
    'authorization',
    'content-type',
    'dnt',
    'origin',
    'user-agent',
    'x-csrftoken',
    'x-requested-with',
]

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'read_data',
    'write_data',
    'rest_framework',
    "corsheaders",
    "metrics_data",
    "intelligent_model"

]

MIDDLEWARE = [
    "corsheaders.middleware.CorsMiddleware",
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

ROOT_URLCONF = 'core.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

WSGI_APPLICATION = 'core.wsgi.application'

# Load from dotenv the database important data

# SAMPLER DATABASE CREDENTIALS
SAMPLER_DB_NAME = os.getenv("SAMPLER_DATABASE_NAME")
SAMPLER_DB_USER = os.getenv("SAMPLER_DATABASE_USER")
SAMPLER_DB_PASSWORD = os.getenv("SAMPLER_DATABASE_PASSWORD")
SAMPLER_DB_HOST = os.getenv("SAMPLER_DATABASE_HOST")
#SAMPLER_DB_PORT = os.getenv("SAMPLER_DATABASE_PORT")

# PROJECT DATABASE CREDENTIALS
PROJECT_DB_NAME = os.getenv("PROJECT_DATABASE_NAME")
PROJECT_DB_USER = os.getenv("PROJECT_DATABASE_USER")
PROJECT_DB_PASSWORD = os.getenv("PROJECT_DATABASE_PASSWORD")
PROJECT_DB_HOST = os.getenv("PROJECT_DATABASE_HOST")


DATABASES = {
    'sampler': {
        'ENGINE': 'mssql',
        'NAME': SAMPLER_DB_NAME,
        'USER': SAMPLER_DB_USER,
        'PASSWORD': SAMPLER_DB_PASSWORD,
        'HOST': SAMPLER_DB_HOST,
        'OPTIONS':{
            'driver': 'ODBC Driver 17 for SQL Server',
           }
    },

    'default': {

        'ENGINE': 'mssql',
        'NAME': PROJECT_DB_NAME,
        'HOST': r'(localdb)\SRLlocal',
        'OPTIONS': {

            'driver': 'ODBC Driver 17 for SQL Server',
            'trusted_connection': 'yes',
        }

    }
}



AUTH_PASSWORD_VALIDATORS = [
    {
        'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator',
    },
    {
        'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator',
    },
    {
        'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator',
    },
    {
        'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator',
    },
]

LANGUAGE_CODE = 'en-us'

TIME_ZONE = 'UTC'

USE_I18N = True

USE_TZ = True

STATIC_URL = 'static/'

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

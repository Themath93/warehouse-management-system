# -*- encoding: utf-8 -*-
"""
Copyright (c) 2019 - present AppSeed.us
"""

import os, environ

env = environ.Env(
    # set casting, default value
    DEBUG=(bool, True)
)

# Build paths inside the project like this: os.path.join(BASE_DIR, ...)
BASE_DIR = os.path.dirname(os.path.dirname(__file__))
CORE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Take environment variables from .env file
environ.Env.read_env(os.path.join(BASE_DIR, '.env'))

# SECURITY WARNING: keep the secret key used in production secret!
SECRET_KEY = env('SECRET_KEY', default='S#perS3crEt_007')

# SECURITY WARNING: don't run with debug turned on in production!
DEBUG = env('DEBUG')

# Assets Management
ASSETS_ROOT = os.getenv('ASSETS_ROOT', '/static/assets') 

# load production server from .env
# ALLOWED_HOSTS        = ['localhost', 'localhost:85', '127.0.0.1',               env('SERVER', default='127.0.0.1') ]
# CSRF_TRUSTED_ORIGINS = ['http://localhost:85', 'http://127.0.0.1', 'https://' + env('SERVER', default='127.0.0.1') ]

ALLOWED_HOSTS        = ['43.200.244.44', '43.200.244.44:85', 'www.mike-de-blog.net',               env('SERVER', default='43.200.244.44') ]
CSRF_TRUSTED_ORIGINS = ['http://43.200.244.44:85', 'http://43.200.244.44', 'https://' + env('SERVER', default='43.200.244.44') ]


# DB
DEFAULT_AUTO_FIELD='django.db.models.AutoField'
# AUTH_USER_MODEL = 'authentication.User'
# Application definition

# Email Setting
EMAIL_BACKEND = 'django.core.mail.backends.smtp.EmailBackend'
EMAIL_PORT = 587
EMAIL_USE_TLS = True
EMAIL_HOST = 'smtp-mail.outlook.com'
EMAIL_HOST_USER = "deyoon@outlook.kr"
EMAIL_HOST_PASSWORD = "as934285qweQWE"


INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'apps.home',  # Enable the inner home (home)
    'apps.authentication',
    'django_prometheus',
]

MIDDLEWARE = [
    # 'django_prometheus.middleware.PrometheusBeforeMiddleware',
    'django.middleware.security.SecurityMiddleware',
    'whitenoise.middleware.WhiteNoiseMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
    'core.middleware.LogMiddleware',
    # 'django_prometheus.middleware.PrometheusAfterMiddleware',
]

ROOT_URLCONF = 'core.urls'
LOGIN_REDIRECT_URL = "home"  # Route defined in home/urls.py
LOGOUT_REDIRECT_URL = "home"  # Route defined in home/urls.py
TEMPLATE_DIR = os.path.join(CORE_DIR, "apps/templates")  # ROOT dir for templates

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [TEMPLATE_DIR],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
                'apps.context_processors.cfg_assets_root',
            ],
        },
    },
]

WSGI_APPLICATION = 'core.wsgi.application'

# Old DB connect oracle cloud with JDBC
DATABASES = {
    'default':{
        'ENGINE':'django.db.backends.oracle',
        'NAME':'fulfill_high', # tnsnames.ora 파일에 등록된 NAME을 등록
        'USER':'web_fulfill', 
        'PASSWORD':'fulfillment123QWE!@#', #Please provide the db password here
    },
    'dw':{
        'ENGINE':'django.db.backends.oracle',
        'NAME':'fulfill_high', # tnsnames.ora 파일에 등록된 NAME을 등록
        'USER':'dw_fulfill', 
        'PASSWORD':'fulfillment123QWE!@#', #Please provide the db password here
    },
    'dm':{
        'ENGINE':'django.db.backends.oracle',
        'NAME':'fulfill_high', # tnsnames.ora 파일에 등록된 NAME을 등록
        'USER':'dm_fulfill', 
        'PASSWORD':'fulfillment123QWE!@#DM', #Please provide the db password here
    }

}

# DB setting for prometheus
# DATABASES = {
#     'default':{
#         'ENGINE':'django.db.backends.oracle',
#         'NAME':'fulfill_high', # tnsnames.ora 파일에 등록된 NAME을 등록
#         'USER':'web_fulfill', 
#         'PASSWORD':'fulfillment123QWE!@#', #Please provide the db password here
#     },
#     'dw':{
#         'ENGINE':'django.db.backends.oracle',
#         'NAME':'fulfill_high', # tnsnames.ora 파일에 등록된 NAME을 등록
#         'USER':'dw_fulfill', 
#         'PASSWORD':'fulfillment123QWE!@#', #Please provide the db password here
#     },
#     'dm':{
#         'ENGINE':'django.db.backends.oracle',
#         'NAME':'fulfill_high', # tnsnames.ora 파일에 등록된 NAME을 등록
#         'USER':'dm_fulfill', 
#         'PASSWORD':'fulfillment123QWE!@#DM', #Please provide the db password here
#     }

# }

# Password validation
# https://docs.djangoproject.com/en/3.0/ref/settings/#auth-password-validators

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

# Internationalization
# https://docs.djangoproject.com/en/3.0/topics/i18n/

LANGUAGE_CODE = 'ko-kr'

TIME_ZONE = 'Asia/Seoul'

USE_I18N = True

USE_L10N = True

USE_TZ = False

#############################################################
# SRC: https://devcenter.heroku.com/articles/django-assets

# Static files (CSS, JavaScript, Images)
# https://docs.djangoproject.com/en/1.9/howto/static-files/
STATIC_ROOT = os.path.join(CORE_DIR, 'staticfiles')
STATIC_URL = '/static/'

# Extra places for collectstatic to find static files.
STATICFILES_DIRS = (
    os.path.join(CORE_DIR, 'apps/static'),
)

#############################################################
#############################################################

# Kafka and ElasticSearch config
KAFKA_SERVERS = os.getenv('KAFKA_SERVERS', '43.205.123.229:9092')
# Generated by Django 5.1.6 on 2025-02-14 11:08

import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('statements', '0003_subject'),
    ]

    operations = [
        migrations.CreateModel(
            name='Teacher',
            fields=[
                ('id', models.AutoField(primary_key=True, serialize=False)),
                ('name', models.CharField(max_length=100, unique=True)),
                ('college', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='statements.college')),
            ],
        ),
    ]

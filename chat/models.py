from django.db import models
from django.contrib.auth.models import User

# class ChatRoom(models.Model):
#     name = models.CharField(max_length=255)
#
#     def __str__(self):
#         return self.name
class Message(models.Model):
    id = models.AutoField(primary_key = True)
    user = models.IntegerField()
    receiver = models.IntegerField()
    content = models.TextField()
    timestamp = models.DateTimeField(auto_now_add=True)


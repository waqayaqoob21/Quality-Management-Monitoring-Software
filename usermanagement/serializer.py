from rest_framework_simplejwt.serializers import TokenObtainPairSerializer
from rest_framework_simplejwt.tokens import RefreshToken


class MyTokenObtainPairSerializer(TokenObtainPairSerializer):
    @classmethod
    def get_token(cls, user):
        # token = super().get_token(user)
        token = RefreshToken.for_user(user)
        # Add custom claims
        token['name'] = user.username
        # Add more custom fields from your custom user model, If you have a
        # custom user model.
        # ...

        return token

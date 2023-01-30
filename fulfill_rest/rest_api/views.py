# rest_api.views.py

from rest_api.models import *
from rest_framework import viewsets
from rest_framework import permissions
from rest_framework.response import Response
from rest_api.serializers import *
# from tutorial.quickstart.serializers import UserSerializer, GroupSerializer


class ShipmentInformationViewSet(viewsets.ModelViewSet):
    """
    API endpoint that allows users to be viewed or edited.
    """
    queryset = ShipmentInformation.objects.all()
    serializer_class = ShipmentInformation
    permission_classes = [permissions.IsAuthenticated]

    def list(self,request):
        page = self.paginate_queryset(self.queryset)
        serializer = self.get_serializer(page, many=True)
        return Response(serializer.data)


# class GroupViewSet(viewsets.ModelViewSet):
#     """
#     API endpoint that allows groups to be viewed or edited.
#     """
#     queryset = Group.objects.all()
#     serializer_class = GroupSerializer
#     permission_classes = [permissions.IsAuthenticated]
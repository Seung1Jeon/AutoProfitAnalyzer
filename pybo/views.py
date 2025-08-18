from django.shortcuts import render

# Create your views here.

def index(request):
    """
    pybo 메인 페이지
    """
    return render(request, 'pybo/index.html')

def test(request):
    """
    pybo 테스트 페이지
    """
    return render(request, 'pybo/test.html')

// AutoProfitAnalyzer JavaScript
document.addEventListener('DOMContentLoaded', function() {
    console.log('AutoProfitAnalyzer 페이지가 로드되었습니다!');
    
    // 환영 메시지 애니메이션
    const welcomeText = document.querySelector('h1');
    if (welcomeText) {
        welcomeText.style.opacity = '0';
        welcomeText.style.transition = 'opacity 1s ease-in';
        setTimeout(() => {
            welcomeText.style.opacity = '1';
        }, 100);
    }
});

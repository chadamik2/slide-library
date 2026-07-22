// Моковые данные для MVP (в реальности приходят по API)
const slidesDatabase = [
    { id: 1, title: "Финансовый отчет Q3", category: "finance", date: "Обновлено: Окт 2023", image: "https://via.placeholder.com/300x160?text=Finance+Report", fileBase64: "" },
    { id: 2, title: "Воронка продаж", category: "marketing", date: "Обновлено: Ноя 2023", image: "https://via.placeholder.com/300x160?text=Sales+Funnel", fileBase64: "" },
    { id: 3, title: "Структура команды", category: "all", date: "Обновлено: Дек 2023", image: "https://via.placeholder.com/300x160?text=Team+Structure", fileBase64: "" }
];

let currentFilter = 'all';
let searchQuery = '';

// Инициализация Office.js
Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        renderSlides();
        document.getElementById('searchInput').addEventListener('input', (e) => {
            searchQuery = e.target.value.toLowerCase();
            renderSlides();
        });
    }
});

// Отрисовка каталога
function renderSlides() {
    const container = document.getElementById('slideContainer');
    container.innerHTML = '';

    const filtered = slidesDatabase.filter(slide => {
        const matchCategory = currentFilter === 'all' || slide.category === currentFilter;
        const matchSearch = slide.title.toLowerCase().includes(searchQuery);
        return matchCategory && matchSearch;
    });

    filtered.forEach(slide => {
        const card = document.createElement('div');
        card.className = 'slide-card';
        card.innerHTML = `
            <img src="${slide.image}" class="slide-preview" alt="${slide.title}">
            <div class="slide-info">
                <p class="slide-title">${slide.title}</p>
                <p class="slide-meta">${slide.date}</p>
                <button class="btn-insert" onclick="insertSlide(${slide.id})">Вставить слайд</button>
            </div>
        `;
        container.appendChild(card);
    });
}

function filterSlides(category) {
    currentFilter = category;
    renderSlides();
}

// Ключевой метод вставки слайда
function insertSlide(id) {
    // Для реального слайда требуется строка base64 формата .pptx
    // В MVP мы имитируем этот процесс. 
    // Метод: context.presentation.insertSlidesFromBase64(base64String);
    
    console.log(`Запрос API на скачивание слайда ID: ${id}`);
    
    // В рамках демо-прототипа вставляем заглушку в виде текста на текущий слайд
    Office.context.document.setSelectedDataAsync(
        `[Успешно загружен слайд: ${slidesDatabase.find(s => s.id === id).title}]`,
        { coercionType: Office.CoercionType.Text },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Ошибка вставки: " + asyncResult.error.message);
            }
        }
    );
}
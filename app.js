// База данных слайдов, привязанная к твоим локальным файлам
const slidesDatabase = [
    { 
        id: 1, 
        title: "Блоки и карточки", 
        category: "structure", 
        date: "Обновлено: 23 Июля 2026", 
        image: "cards.jpg", // Твоя картинка-превью
        fileUrl: "cards.pptx" // Твой реальный файл
    },
    { 
        id: 2, 
        title: "Схемы и диаграммы", 
        category: "graphics", 
        date: "Обновлено: 23 Июля 2026", 
        image: "diagrams.jpg", 
        fileUrl: "diagrams.pptx" 
    },
    { 
        id: 3, 
        title: "Дорожная карта (Roadmap)", 
        category: "planning", 
        date: "Обновлено: 23 Июля 2026", 
        image: "roadmap.jpg", 
        fileUrl: "roadmap.pptx" 
    },
    { 
        id: 4, 
        title: "Таблица данных", 
        category: "data", 
        date: "Обновлено: 23 Июля 2026", 
        image: "table.jpg", 
        fileUrl: "table.pptx" 
    }
];

let currentFilter = 'all';
let searchQuery = '';

// Инициализация Office.js
Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        renderSlides();
        
        // Слушатель для строки поиска
        const searchInput = document.getElementById('searchInput');
        if (searchInput) {
            searchInput.addEventListener('input', (e) => {
                searchQuery = e.target.value.toLowerCase();
                renderSlides();
            });
        }
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

    if (filtered.length === 0) {
        container.innerHTML = '<p style="color: #666; text-align: center;">Слайды не найдены</p>';
        return;
    }

    filtered.forEach(slide => {
        const card = document.createElement('div');
        card.className = 'slide-card';
        card.innerHTML = `
            <img src="${slide.image}" class="slide-preview" alt="${slide.title}" onerror="this.src='https://via.placeholder.com/300x160?text=Превью+не+найдено'">
            <div class="slide-info">
                <p class="slide-title">${slide.title}</p>
                <p class="slide-meta">${slide.date}</p>
                <button class="btn-insert" id="btn-${slide.id}" onclick="insertRealSlide('${slide.fileUrl}', 'btn-${slide.id}')">Вставить слайд</button>
            </div>
        `;
        container.appendChild(card);
    });
}

// Функция фильтрации
window.filterSlides = function(category) {
    currentFilter = category;
    renderSlides();
}

// Логика скачивания и вставки реального файла .pptx
async function insertRealSlide(fileUrl, buttonId) {
    const btn = document.getElementById(buttonId);
    const originalText = btn.innerText;
    
    try {
        // Визуальный фидбек для пользователя
        btn.innerText = "Загрузка...";
        btn.style.backgroundColor = "#ccc";
        btn.disabled = true;

        // 1. Скачиваем файл .pptx с нашего локального сервера
        const response = await fetch(fileUrl);
        if (!response.ok) throw new Error("Файл не найден на сервере");
        const blob = await response.blob();

        // 2. Конвертируем в Base64
        const reader = new FileReader();
        reader.readAsDataURL(blob);
        reader.onloadend = async function () {
            const base64String = reader.result.toString().split(',')[1];

            // 3. Обращаемся к API PowerPoint для вставки
            await PowerPoint.run(async (context) => {
                context.presentation.insertSlidesFromBase64(base64String, {
                    formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme 
                });
                await context.sync();
            });

            // Возвращаем кнопку в исходное состояние
            btn.innerText = "Успешно!";
            btn.style.backgroundColor = "#107c41";
            
            setTimeout(() => {
                btn.innerText = originalText;
                btn.style.backgroundColor = "";
                btn.disabled = false;
            }, 2000);
        };
    } catch (error) {
        console.error("Ошибка при вставке слайда:", error);
        btn.innerText = "Ошибка";
        btn.style.backgroundColor = "#d83b01";
        
        setTimeout(() => {
            btn.innerText = originalText;
            btn.style.backgroundColor = "";
            btn.disabled = false;
        }, 2000);
    }
}
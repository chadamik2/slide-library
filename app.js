// Моковые данные для MVP. 
// Вместо ссылки на файл здесь указано локальное имя файла, который лежит в нашей папке.
const slidesDatabase = [
    { id: 1, title: "Финансовый отчет Q3", category: "finance", date: "Обновлено: Окт 2023", image: "https://via.placeholder.com/300x160/0078D4/FFFFFF?text=Финансы", fileUrl: "template.pptx" },
    { id: 3, title: "Структура команды", category: "all", date: "Обновлено: Дек 2023", image: "https://via.placeholder.com/300x160/D83B01/FFFFFF?text=Команда", fileUrl: "template.pptx" }
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
                <!-- Передаем URL файла в функцию вставки -->
                <button class="btn-insert" onclick="insertRealSlide('${slide.fileUrl}')">Вставить слайд</button>
            </div>
        `;
        container.appendChild(card);
    });
}

function filterSlides(category) {
    currentFilter = category;
    renderSlides();
}

// ---------------------------------------------------------
// НОВАЯ ЛОГИКА ВСТАВКИ РЕАЛЬНОГО СЛАЙДА
// ---------------------------------------------------------

async function insertRealSlide(fileUrl) {
    try {
        // 1. Изменяем текст кнопки, чтобы показать процесс загрузки (для хорошего UX)
        console.log(`Скачиваем файл: ${fileUrl}`);
        
        // 2. Скачиваем файл .pptx
        const response = await fetch(fileUrl);
        if (!response.ok) throw new Error("Не удалось загрузить файл шаблона");
        const blob = await response.blob();

        // 3. Конвертируем скачанный файл в формат Base64
        const reader = new FileReader();
        reader.readAsDataURL(blob);
        reader.onloadend = async function () {
            // Убираем приставку "data:application/vnd.openxmlformats...;base64," 
            // Оставляем только чистый base64 код
            const base64String = reader.result.toString().split(',')[1];

            // 4. Вставляем слайд через API PowerPoint
            await PowerPoint.run(async (context) => {
                context.presentation.insertSlidesFromBase64(base64String, {
                    // Используем тему текущей презентации (чтобы цвета и шрифты подстроились)
                    formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme 
                });
                await context.sync();
                console.log("Слайд успешно вставлен!");
            });
        };
    } catch (error) {
        console.error("Ошибка при вставке слайда:", error);
        // В случае ошибки выводим стандартный диалог Office
        Office.context.ui.displayDialogAsync('https://localhost:3000/error.html', { height: 30, width: 20 });
    }
}
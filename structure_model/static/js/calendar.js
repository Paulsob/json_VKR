// globals for currently displayed month/year and selected route state
let currentYear, currentMonth;
let selectedRoute = "55";
let lastLoadedDay = null;

function initCalendar() {
  const now = new Date();
  currentYear = now.getFullYear();
  currentMonth = now.getMonth();

  const routeSelect = document.getElementById("routeSelect");
  if (routeSelect) {
    // Пока грузим — блокируем селект и оставляем "Загрузка..."
    routeSelect.disabled = true;

    // если в html был установлен какой-то value — используем его как предпочтение
    selectedRoute = routeSelect.value || selectedRoute || "55";

    // Загрузка списка маршрутов с бекенда
    fetch("/api/routes")
      .then((response) => {
        if (!response.ok) throw new Error("Ошибка при получении маршрутов");
        return response.json();
      })
      .then((routes) => {
        // очищаем селект и заполняем
        routeSelect.innerHTML = "";

        if (!Array.isArray(routes) || routes.length === 0) {
          const opt = document.createElement("option");
          opt.value = "";
          opt.disabled = true;
          opt.selected = true;
          opt.textContent = "Нет маршрутов";
          routeSelect.appendChild(opt);
          routeSelect.disabled = true;
        } else {
          routes.forEach((rn) => {
            const opt = document.createElement("option");
            opt.value = String(rn);
            opt.textContent = `Маршрут ${rn}`;
            // выбираем, если совпадает с выбранным (например 55)
            if (String(rn) === String(selectedRoute)) opt.selected = true;
            routeSelect.appendChild(opt);
          });

          // если ни один не был выбран, выбираем первый
          if (!routeSelect.value) {
            routeSelect.value = String(routes[0]);
          }
          selectedRoute = routeSelect.value;
          routeSelect.disabled = false;
        }

        // слушатель изменения маршрута
        routeSelect.addEventListener("change", () => {
          selectedRoute = routeSelect.value;
          if (lastLoadedDay) {
            loadSchedule(lastLoadedDay, true);
          }
        });

        // После загрузки маршрутов загружаем календарь
        return fetch("/calendar-data");
      })
      .then((response) => {
        if (!response) return; // если предыдущ шаг уже упал
        if (!response.ok) throw new Error("Ошибка при получении календаря");
        return response.json();
      })
      .then((daysWithSchedules) => {
        if (Array.isArray(daysWithSchedules)) {
          renderCalendar(currentYear, currentMonth, daysWithSchedules);
        } else {
          renderCalendar(currentYear, currentMonth, []);
        }
      })
      .catch((err) => {
        console.error("Ошибка инициализации календаря/маршрутов:", err);
        // В случае ошибки с маршрутами — оставляем понятное сообщение в селекте
        if (routeSelect) {
          routeSelect.innerHTML = "";
          const opt = document.createElement("option");
          opt.value = "";
          opt.disabled = true;
          opt.selected = true;
          opt.textContent = "Ошибка загрузки маршрутов";
          routeSelect.appendChild(opt);
          routeSelect.disabled = true;
        }
        // пробуем всё равно загрузить календарь (на случай, если routes упали, но calendar-data — ок)
        fetch("/calendar-data")
          .then((r) => (r.ok ? r.json() : []))
          .then((daysWithSchedules) => {
            renderCalendar(currentYear, currentMonth, daysWithSchedules || []);
          })
          .catch((e) => {
            console.error("Не удалось загрузить календарь:", e);
            renderCalendar(currentYear, currentMonth, []);
          });
      });
  } else {
    // если селекта нет — просто загружаем календарь
    fetch("/calendar-data")
      .then((response) => response.json())
      .then((daysWithSchedules) => {
        renderCalendar(currentYear, currentMonth, daysWithSchedules);
      })
      .catch((err) => {
        console.error("Не удалось загрузить данные календаря:", err);
      });
  }
}

function renderCalendar(year, month, daysWithSchedules) {
  // сохраняем текущие отображаемые год и месяц
  currentYear = year;
  currentMonth = month;

  const container = document.getElementById("calendar-container");
  const firstDay = new Date(year, month, 1);
  const lastDay = new Date(year, month + 1, 0);
  const daysInMonth = lastDay.getDate();

  let html = `
        <div class="d-flex justify-content-between align-items-center mb-3">
            <h4>${getMonthName(month)} ${year}</h4>
            <div>
                <button class="btn btn-sm btn-outline-secondary" id="prevMonth">&lt;</button>
                <button class="btn btn-sm btn-outline-secondary" id="nextMonth">&gt;</button>
            </div>
        </div>
        <div class="table-responsive">
            <table class="table table-bordered">
                <thead class="table-light">
                    <tr>
                        <th>Пн</th><th>Вт</th><th>Ср</th><th>Чт</th><th>Пт</th><th>Сб</th><th>Вс</th>
                    </tr>
                </thead>
                <tbody>
    `;

  // Определяем, с какого дня недели начинается месяц (понедельник=1 ... воскресенье=7)
  let startDayOfWeek = firstDay.getDay();
  if (startDayOfWeek === 0) startDayOfWeek = 7;

  // Пустые ячейки до начала месяца
  for (let i = 1; i < startDayOfWeek; i++) {
    html += "<td></td>";
  }

  // Дни месяца
  let dayOfWeek = startDayOfWeek;
  for (let day = 1; day <= daysInMonth; day++) {
    if (dayOfWeek === 1) html += "<tr>";

    const hasSchedule =
      Array.isArray(daysWithSchedules) && daysWithSchedules.includes(day);
    const today = new Date();
    const isToday =
      day === today.getDate() &&
      month === today.getMonth() &&
      year === today.getFullYear();

    html += `
            <td class="${hasSchedule ? "table-success" : ""} ${
      isToday ? "table-primary" : ""
    }">
                <div class="day-number">${day}</div>
                ${
                  hasSchedule
                    ? `<button class="btn btn-sm btn-outline-success w-100 mt-1 schedule-btn"
                            data-day="${day}">
                        Расписание
                    </button>`
                    : `<span class="text-muted small">нет данных</span>`
                }
            </td>
        `;

    if (dayOfWeek === 7) {
      html += "</tr>";
      dayOfWeek = 0;
    }
    dayOfWeek++;
  }

  // Пустые ячейки в конце месяца
  if (dayOfWeek !== 1) {
    for (let i = dayOfWeek; i <= 7; i++) {
      html += "<td></td>";
    }
    html += "</tr>";
  }

  html += `
                </tbody>
            </table>
        </div>
    `;

  container.innerHTML = html;

  // один обработчик для всех кнопок
  document.querySelectorAll(".schedule-btn").forEach((button) => {
    button.addEventListener("click", function () {
      const day = this.getAttribute("data-day");
      requestRecalculateIfNeeded(day).finally(() => {
        loadSchedule(day);
      });
    });
  });

  document
    .getElementById("prevMonth")
    .addEventListener("click", () => changeMonth(-1));
  document
    .getElementById("nextMonth")
    .addEventListener("click", () => changeMonth(1));
}

function changeMonth(delta) {
  let newMonth = currentMonth + delta;
  let newYear = currentYear;

  if (newMonth < 0) {
    newMonth = 11;
    newYear--;
  } else if (newMonth > 11) {
    newMonth = 0;
    newYear++;
  }

  // Загружаем дни с расписаниями для нового месяца (если ваш backend поддерживает параметры, лучше передавать month/year)
  fetch("/calendar-data")
    .then((response) => response.json())
    .then((daysWithSchedules) => {
      renderCalendar(newYear, newMonth, daysWithSchedules);
    })
    .catch((err) => {
      console.error("Не удалось загрузить календарь для нового месяца:", err);
    });
}

function getMonthName(monthIndex) {
  const months = [
    "Январь",
    "Февраль",
    "Март",
    "Апрель",
    "Май",
    "Июнь",
    "Июль",
    "Август",
    "Сентябрь",
    "Октябрь",
    "Ноябрь",
    "Декабрь",
  ];
  return months[monthIndex];
}

function loadSchedule(day, skipLoadingMessage) {
  const display = document.getElementById("scheduleDisplay");
  if (!display) return;

const routeSelect = document.getElementById("routeSelect");
if (routeSelect) {
  selectedRoute = routeSelect.value || "55";
}


  if (!skipLoadingMessage) {
    display.innerHTML = "<p>Загрузка расписания...</p>";
  }

  lastLoadedDay = parseInt(day, 10);
  updateRouteHint();

  const params = new URLSearchParams({ route: selectedRoute });

  fetch(`/api/schedule/${day}?${params.toString()}`)
    .then((response) => response.json())
    .then((data) => {
      if (!data || !data.success) {
        const errMsg =
          data && data.error
            ? data.error
            : "Не удалось получить данные расписания";
        display.innerHTML = `<div class="alert alert-warning">${errMsg}</div>`;
        return;
      }

      let rows = Array.isArray(data.rows) ? data.rows : [];

      //------------------------------------------------------------------
      // 2) УДАЛЯЕМ каждую вторую строку (строки часов работы)
      //    Остаются строки: 0, 2, 4, 6…
      //------------------------------------------------------------------
      rows = rows.filter((row, index) => index % 2 !== 0);

      //------------------------------------------------------------------
      // 3) Определяем, какие колонки оставить
      //------------------------------------------------------------------
      const colCount = rows.length
        ? Math.max(...rows.map((r) => (Array.isArray(r) ? r.length : 0)))
        : 0;
      const keepCols = [];

      for (let c = 0; c < colCount; c++) {
        let any = false;
        for (let r = 0; r < rows.length; r++) {
          const cell = rows[r] && rows[r][c];
          if (
            cell !== null &&
            cell !== undefined &&
            String(cell).trim() !== ""
          ) {
            any = true;
            break;
          }
        }
        if (any) keepCols.push(c);
      }

      if (keepCols.length === 0 && colCount > 0) {
        for (let i = 0; i < colCount; i++) keepCols.push(i);
      }

      // заголовки столбцов: если сервер прислал data.columns — используем их, иначе берём стандартные названия
      const realIndexes = {
        0: "Номер маршрута",
        5: "Отправление 1 смена",
        6: "Прибытие 1 смена",
        10: "Отправление 2 смена",
        11: "Прибытие 2 смена",
      };

      let headers = keepCols.map((i) => realIndexes[i] || `Поле ${i + 1}`);

      // вычисляем день недели по текущему отображаемому месяцу/году
      const d = new Date(currentYear, currentMonth, parseInt(day, 10));
      const weekdays = [
        "Воскресенье",
        "Понедельник",
        "Вторник",
        "Среда",
        "Четверг",
        "Пятница",
        "Суббота",
      ];
      const weekdayName = weekdays[d.getDay()];
      const scheduleType = data.is_weekend ? "Выходной день" : "Рабочий день";
      const currentRoute = data.route || selectedRoute;

      // Сборка таблицы: заголовок в шапке таблицы + строка заголовков колонок
      let html = `
<style>
.schedule-table td, .schedule-table th {
    white-space: nowrap;
    vertical-align: middle;
    padding: 0.4rem;
}
</style>
<div class="table-responsive">
    <table class="table table-bordered schedule-table">
        <thead>
            <tr>
                <th colspan="${
                  headers.length
                }" class="text-center align-middle">
                    Расписание — маршрут ${currentRoute}, ${scheduleType.toLowerCase()} — ${weekdayName}, ${day} ${getMonthName(
        currentMonth
      )} ${currentYear}
                </th>
            </tr>
            <tr>
                ${headers
                  .map(
                    (h) =>
                      `<th>${String(h)
                        .replace(/</g, "&lt;")
                        .replace(/>/g, "&gt;")}</th>`
                  )
                  .join("")}
            </tr>
        </thead>
        <tbody>
`;

      // строки таблицы — только выбранные колонки
      rows.forEach((row) => {
        html += "<tr>";
        keepCols.forEach((c) => {
          let cell =
            row && row[c] !== undefined && row[c] !== null && row[c] !== ""
              ? String(row[c])
              : "&nbsp;";
          cell = cell.replace(/</g, "&lt;").replace(/>/g, "&gt;");
          html += `<td>${cell}</td>`;
        });
        html += "</tr>";
      });

      html += `
        </tbody>
    </table>
</div>
`;

      // Если сервер прислал какую-то дополнительную информацию о маршруте — показываем над таблицей (или можно убрать)
      if (data.route_info) {
        html =
          `<p class="fw-bold text-primary">${String(data.route_info).replace(
            /</g,
            "&lt;"
          )}</p>` + html;
      }

      const driverSummary = buildDriverSummary(data.drivers);
      display.innerHTML = html + driverSummary;
      updateRouteHint(scheduleType, currentRoute);
    })
    .catch((err) => {
      console.error("Ошибка:", err);
      display.innerHTML =
        '<div class="alert alert-danger">Не удалось загрузить расписание.</div>';
    });
}

function requestRecalculateIfNeeded(day) {
  const weekendSwitch = document.getElementById("weekendSwitch");
  if (!weekendSwitch || !weekendSwitch.checked) {
    return Promise.resolve();
  }

  const body = JSON.stringify({
    allow_weekend: true,
    route: selectedRoute,
  });

  return fetch(`/api/recalculate/${day}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body,
  })
    .then((response) => {
      if (!response.ok) {
        console.error("Ошибка пересчета расписания:", response.statusText);
      }
    })
    .catch((err) => {
      console.error("Не удалось пересчитать расписание:", err);
    });
}

function updateRouteHint(scheduleType, routeNumber) {
  const hintEl = document.getElementById("routeHint");
  if (!hintEl) return;

  if (routeNumber && scheduleType) {
    hintEl.textContent = `Показано расписание маршрута ${routeNumber} (${scheduleType.toLowerCase()}).`;
    return;
  }

  const baseRoute = document.getElementById("routeSelect")
    ? document.getElementById("routeSelect").value
    : selectedRoute;
  if (lastLoadedDay) {
    hintEl.textContent = `Маршрут ${baseRoute} выбран, доступно расписание за ${lastLoadedDay}-е число.`;
  } else {
    hintEl.textContent = `Выберите маршрут ${baseRoute} и нажмите «Расписание» в календаре.`;
  }
}

function buildDriverSummary(drivers) {
  if (!Array.isArray(drivers) || !drivers.length) {
    return "";
  }

  const rows = drivers
    .map((entry, index) => {
      const shift = entry.shift || "—";
      const tabNo = escapeHtml(entry.tab_no || "—");
      const graph = escapeHtml(entry.graph_type || "—");
      return `
            <tr>
                <td>${index + 1}</td>
                <td>${shift}</td>
                <td>${tabNo}</td>
                <td>${graph}</td>
            </tr>
        `;
    })
    .join("");

  return `
        <div class="card mt-3">
            <div class="card-header bg-light">
                <strong>Водители на линии</strong>
            </div>
            <div class="card-body p-0">
                <div class="table-responsive mb-0">
                    <table class="table table-sm table-striped mb-0">
                        <thead class="table-light">
                            <tr>
                                <th>#</th>
                                <th>Смена</th>
                                <th>Таб. №</th>
                                <th>График</th>
                            </tr>
                        </thead>
                        <tbody>${rows}</tbody>
                    </table>
                </div>
            </div>
        </div>
    `;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

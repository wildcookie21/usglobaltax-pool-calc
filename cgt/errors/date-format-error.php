<div style="font-family: Arial, sans-serif; padding: 24px; max-width: 1200px; margin: 0 auto; color: #222;">

  <div style="background: #fff3f3; border: 1px solid #f1b8b8; border-radius: 10px; padding: 18px 20px; margin-bottom: 24px;">
    <h2 style="margin: 0 0 12px; color: #b00020;">Ошибка обработки даты</h2>

    <p style="margin: 6px 0;">Exel ячейка: <b>А <?= htmlspecialchars($row) ?></b></p>

    <?php if (!empty($formatCode)): ?>
      <p style="margin: 6px 0;">Формат ячейки: <b><?= htmlspecialchars($formatCode) ?></b></p>
    <?php endif; ?>
  </div>

  <p style="font-size: 16px; line-height: 1.5; margin-bottom: 24px;">
    Проверьте формат ячейки с датой в Excel. Ниже показана настройка формата даты.
  </p>

  <div style="display: flex; flex-wrap: wrap; gap: 24px; align-items: flex-start;">

    <div style="flex: 1 1 420px; background: #f8f9fb; border: 1px solid #ddd; border-radius: 10px; padding: 16px;">
      <h3 style="margin-top: 0;">Шаг 1</h3>

      <p style="line-height: 1.5;">
        Левой кнопкой мыши выделите <b>столбец с датами</b>. Нажмите правой кнопкой мыши и выберите <b>«Формат ячеек…».</b>
      </p>

      <img src="/images/date-format-help-1.jpg"
        alt="Шаг 1"
        style="width: 100%; max-width: 100%; border: 1px solid #ccc; border-radius: 8px;">
    </div>

    <div style="flex: 1 1 420px; background: #f8f9fb; border: 1px solid #ddd; border-radius: 10px; padding: 16px;">
      <h3 style="margin-top: 0;">Шаг 2</h3>

      <p style="line-height: 1.5;">
        В появившемся окне выберите раздел <b>«Дата»</b>, выберите формат <b>14.03.12</b>, нажмите <b>«ОК»</b>.
      </p>

      <img src="/images/date-format-help-2.jpg"
        alt="Шаг 2"
        style="width: 100%; max-width: 100%; border: 1px solid #ccc; border-radius: 8px;">
    </div>

  </div>

  <div style="margin-top: 18px;">
    <a href="/"
      style="display: inline-block; padding: 10px 16px; background: #1f6feb; color: #fff; text-decoration: none; border-radius: 8px;">
      Вернуться на главную страницу
    </a>
  </div>

</div>
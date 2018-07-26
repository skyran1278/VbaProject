clc; clear; close all;

maxEQ = 30;
positiveDL = 20 / 3;
negativeDL = 20 / 3 * 2;

beamLength = 10;

x = 0 : 0.01 : beamLength;
x_length = length(x);

midline = zeros(1, x_length);

EQ = -maxEQ + 2 * maxEQ / beamLength * x;
NEQ = -EQ;
DL = -positiveDL + 4 * (positiveDL + negativeDL) * (x / beamLength - 0.5) .^ 2;

% DL + EQ
negativeMn = EQ .* (EQ >= 0) + NEQ .* (NEQ >= 0) + 1.0 * DL .* (DL >= 0);
positiveMn = EQ .* (EQ <= 0) + NEQ .* (NEQ <= 0) + 1.0 * DL .* (DL <= 0);

topLeftRebar = max(negativeMn(x <= beamLength / 3));
topRightRebar = max(negativeMn(x >= beamLength / 3));

% 負彎矩需求曲線
topRebar = [topLeftRebar - topLeftRebar / (beamLength / 2) * x(x <= beamLength / 2), topRightRebar / (beamLength / 2) * x(x > beamLength / 2) - topRightRebar];

% 取正
botLeftRebar = -min(positiveMn(x <= beamLength / 3));
botMidRebar = -min(positiveMn(x >= 1 * beamLength / 4 & x <= 3 * beamLength / 4));
botRightRebar = -min(positiveMn(x >= 2 * beamLength / 3));

botRebarDL = 4 * botMidRebar * (x / beamLength - 0.5) .^ 2 - botMidRebar;

botRebar = min([EQ; NEQ; botRebarDL]);
botRebarOtherMethod = [-botLeftRebar + (botLeftRebar - botMidRebar) / (beamLength / 2) * x(x <= beamLength / 2), -botMidRebar + (botMidRebar - botRightRebar) / (beamLength / 2) * (x(x > beamLength / 2) - (beamLength / 2)) ];

% bot = - botMidRebar * ones(1, x_length);

greenColor = [26 188 156] / 256;
blueColor = [52 152 219] / 256;
redColor = [233 88 73] / 256;
grayColor = [0.5 0.5 0.5];

% 重力、地震力的實際需求
figure;
plot(x, midline, 'Color', grayColor, 'LineWidth', 1.5);
hold on;
plot(x, EQ, '-', 'Color', redColor, 'LineWidth', 1.5);
legendEQ = plot(x, NEQ, '-', 'Color', redColor, 'LineWidth', 1.5);
legendGravity = plot(x, DL, '--', 'Color', redColor, 'LineWidth', 1.5);
axis([0 beamLength -50 50]);
legend([legendEQ, legendGravity], 'EQ', 'Gravity', 'Location', 'northeast');
title('Mn');
xlabel('m');
ylabel('tf-m');

% 線性疊加
figure;
plot(x, midline, 'Color', grayColor, 'LineWidth', 1.5);
hold on;
plot(x, EQ, 'Color', grayColor, 'LineWidth', 1.5);
legendEQ = plot(x, NEQ, 'Color', grayColor, 'LineWidth', 1.5);
legendGravity = plot(x, DL, 'Color', grayColor, 'LineWidth', 1.5);
legendMn = plot(x, positiveMn, 'Color', redColor, 'LineWidth', 1.5);
plot(x, negativeMn, 'Color', redColor, 'LineWidth', 1.5);
axis([0 beamLength -50 50]);
legend([legendEQ, legendGravity, legendMn], 'EQ', 'Gravity', 'Linear Add', 'Location', 'northeast');
title('Mn');
xlabel('m');
ylabel('tf-m');

% 左端為 0~1/3 的最大值
% 中央是 1/4~3/4 的最大值
% 右端是 2/3~1 的最大值
figure;
plot(x, midline, '-k');
hold on;
plot(x, EQ, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, NEQ, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, DL, 'Color', grayColor, 'LineWidth', 1.5);
legendMn = plot(x, positiveMn, 'Color', redColor, 'LineWidth', 1.5);
plot(x, negativeMn, 'Color', redColor, 'LineWidth', 1.5);
legendActural = plot(0, topLeftRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength / 2, 0, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength, topRightRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(0, -botLeftRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength / 2, -botMidRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength, -botRightRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
axis([0 beamLength -50 50]);
legend([legendActural, legendMn(1)], 'Actural Rebar', 'Demand', 'Location', 'northeast');
title('Mn');
xlabel('m');
ylabel('tf-m');

% 實際配筋應該會再大一點
% 而且最少會有兩支的限制
% 這裡取最 critical 的情況
figure;
plot(x, midline, '-k');
hold on;
plot(x, EQ, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, NEQ, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, DL, 'Color', grayColor, 'LineWidth', 1.5);
legendMn = plot(x, positiveMn, 'Color', redColor, 'LineWidth', 1.5);
plot(x, negativeMn, 'Color', redColor, 'LineWidth', 1.5);
legendActural = plot(0, topLeftRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength / 2, 10, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength, topRightRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(0, -botLeftRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength / 2, -botMidRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength, -botRightRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
axis([0 beamLength -50 50]);
legend([legendActural, legendMn(1)], 'Actural Rebar', 'Demand', 'Location', 'northeast');
title('Mn');
xlabel('m');
ylabel('tf-m');

% 再來就是依據現有配筋依據需求曲線做優化
% 首先是上層筋的部分
% 中間沒有需求
% 兩端主要由耐震控制
% 我們就直接拉直
figure;
plot(x, midline, '-k');
hold on;
plot(x, EQ, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, NEQ, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, DL, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, positiveMn, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, negativeMn, 'Color', redColor, 'LineWidth', 1.5);
multiRebar = plot(x, topRebar, 'Color', blueColor, 'LineWidth', 1.5);
% plot(x, botRebar, 'Color', blueColor, 'LineWidth', 1.5);
legendActural = plot(0, topLeftRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength / 2, 10, 'o', 'Color', grayColor, 'LineWidth', 1.5);
plot(beamLength, topRightRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(0, -botLeftRebar, 'o', 'Color', grayColor, 'LineWidth', 1.5);
plot(beamLength / 2, -botMidRebar, 'o', 'Color', grayColor, 'LineWidth', 1.5);
plot(beamLength, -botRightRebar, 'o', 'Color', grayColor, 'LineWidth', 1.5);
axis([0 beamLength -50 50]);
legend([legendActural, multiRebar], 'Actural Rebar', 'Multi Rebar', 'Location', 'southeast');
title('Mn');
xlabel('m');
ylabel('tf-m');

% 接下來是下層筋的部分
% 這部分就比較複雜了
% 左右兩端由耐震控制
% 中央我們原本預估是由重力控制
% 後來真的下去做的時候發現會有地震力的因素參雜進來了
% 如果中間依照重力，兩端依據地震力取大值會如右方藍色的線
figure;
plot(x, midline, '-k');
hold on;
plot(x, EQ, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, NEQ, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, DL, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, positiveMn, 'Color', redColor, 'LineWidth', 1.5);
plot(x, negativeMn, 'Color', grayColor, 'LineWidth', 1.5);
multiRebar = plot(x, botRebar, 'Color', blueColor, 'LineWidth', 1.5);
legendActural = plot(0, topLeftRebar, 'o', 'Color', grayColor, 'LineWidth', 1.5);
plot(beamLength / 2, 10, 'o', 'Color', grayColor, 'LineWidth', 1.5);
plot(beamLength, topRightRebar, 'o', 'Color', grayColor, 'LineWidth', 1.5);
plot(0, -botLeftRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength / 2, -botMidRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength, -botRightRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
axis([0 beamLength -50 50]);
legend([legendActural, multiRebar], 'Actural Rebar', 'Multi Rebar', 'Location', 'northeast');
title('Mn');
xlabel('m');
ylabel('tf-m');

% 可以發現綠色圈圈的部分多估了
% 紅色部分的不保守
figure;
plot(x, midline, '-k');
hold on;
plot(x, EQ, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, NEQ, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, DL, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, positiveMn, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, negativeMn, 'Color', grayColor, 'LineWidth', 1.5);
multiRebar = plot(x, botRebar, 'Color', grayColor, 'LineWidth', 1.5);
legendActural = plot(0, topLeftRebar, 'o', 'Color', grayColor, 'LineWidth', 1.5);
plot(beamLength / 2, 10, 'o', 'Color', grayColor, 'LineWidth', 1.5);
plot(beamLength, topRightRebar, 'o', 'Color', grayColor, 'LineWidth', 1.5);
plot(0, -botLeftRebar, 'o', 'Color', grayColor, 'LineWidth', 1.5);
plot(beamLength / 2, -botMidRebar, 'o', 'Color', grayColor, 'LineWidth', 1.5);
plot(beamLength, -botRightRebar, 'o', 'Color', grayColor, 'LineWidth', 1.5);
fill(x .* (positiveMn - botRebar > 0), (positiveMn - botRebar) .* (positiveMn - botRebar > 0), greenColor, 'edgeColor', 'none')
axis([0 beamLength -50 50]);
legend([legendActural, multiRebar], 'Actural Rebar', 'Multi Rebar', 'Location', 'northeast');
title('Mn');
xlabel('m');
ylabel('tf-m');

% 而如果我們直接拉直線
% 效益會下降
% 從整體 8.5% 的效益下降到 5.5%
figure;
plot(x, midline, '-k');
hold on;
plot(x, EQ, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, NEQ, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, DL, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, positiveMn, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, negativeMn, 'Color', grayColor, 'LineWidth', 1.5);
multiRebar = plot(x, botRebarOtherMethod, 'Color', blueColor, 'LineWidth', 1.5);
legendActural = plot(0, topLeftRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength / 2, 10, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength, topRightRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(0, -botLeftRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength / 2, -botMidRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength, -botRightRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
axis([0 beamLength -50 50]);
legend([legendActural, multiRebar], 'Actural Rebar', 'Multi Rebar', 'Location', 'northeast');
title('Mn');
xlabel('m');
ylabel('tf-m');

% 所以如果想要從現有配筋就有很好的效果的話 ( 如果已經產生配筋表格了)
% 那我們就會需要更多的資料
figure;
plot(x, midline, '-k');
hold on;
plot(x, EQ, 'Color', grayColor, 'LineWidth', 1.5);
legendEQ = plot(x, NEQ, 'Color', grayColor, 'LineWidth', 1.5);
legendGravity = plot(x, DL, 'Color', grayColor, 'LineWidth', 1.5);
legendMn = plot(x, positiveMn, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, negativeMn, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, topRebar, 'Color', blueColor, 'LineWidth', 1.5);
multiRebar = plot(x, botRebarOtherMethod, 'Color', blueColor, 'LineWidth', 1.5);
legendActural = plot(0, topLeftRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength / 2, 10, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength, topRightRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(0, -botLeftRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength / 2, -botMidRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength, -botRightRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
axis([0 beamLength -50 50]);
legend([legendActural, multiRebar], 'Actural Rebar', 'Multi Rebar', 'Location', 'northeast');
title('Mn');
xlabel('m');
ylabel('tf-m');

% 所以如果想要從現有配筋就有很好的效果的話 ( 如果已經產生配筋表格了)
% 那我們就會需要更多的資料
figure;
plot(x, midline, '-k');
hold on;
plot(x, EQ, 'Color', grayColor, 'LineWidth', 1.5);
legendEQ = plot(x, NEQ, 'Color', grayColor, 'LineWidth', 1.5);
legendGravity = plot(x, DL, 'Color', redColor, 'LineWidth', 1.5);
legendMn = plot(x, positiveMn, 'Color', grayColor, 'LineWidth', 1.5);
plot(x, negativeMn, 'Color', grayColor, 'LineWidth', 1.5);
legendActural = plot(0, topLeftRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength / 2, 10, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength, topRightRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(0, -botLeftRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength / 2, -botMidRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
plot(beamLength, -botRightRebar, 'o', 'Color', redColor, 'LineWidth', 1.5);
axis([0 beamLength -50 50]);
legend([legendActural, legendGravity], 'Actural Rebar', 'Gravity', 'Location', 'northeast');
title('Mn');
xlabel('m');
ylabel('tf-m');

% figure;
% plot(x, midline, '-k');
% hold on;
% legendEQ = plot(x, EQ, '-k', x, NEQ, '-k');
% legendGravity = plot(x, DL, '-k');
% legendMn = plot(x, negativeMn, '-r', x, positiveMn, '-r');
% % % plot(x, botRebarDL, '--b');
% % % plot(x, bot, '--g');
% % plot(x, topRebar, '-b');
% % plot(x, botRebar, '-b');
% % plot(x, botRebarOtherMethod, '-g');

% plot(0, topLeftRebar, 'or');
% plot(beamLength / 2, 0, 'or');
% plot(beamLength, topRightRebar, 'or');
% plot(0, -botLeftRebar, 'or');
% plot(beamLength / 2, -botMidRebar, 'or');
% plot(beamLength, -botRightRebar, 'or');
% axis([0 beamLength -50 50]);
% legend([legendEQ(1), legendGravity, legendMn(1)], 'EQ', 'Gravity', 'Linear Add', 'Location', 'northeast');
% title('Mn');
% xlabel('m');
% ylabel('tf-m');

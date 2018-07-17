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
negativeMoment = EQ .* (EQ >= 0) + NEQ .* (NEQ >= 0) + 1.4 * DL .* (DL >= 0);
positiveMoment = EQ .* (EQ <= 0) + NEQ .* (NEQ <= 0) + 1.4 * DL .* (DL <= 0);

topLeftRebar = max(negativeMoment(x <= beamLength / 3));
topRightRebar = max(negativeMoment(x >= beamLength / 3));

% 負彎矩需求曲線
topRebar = [topLeftRebar - topLeftRebar / (beamLength / 2) * x(x <= beamLength / 2), topRightRebar / (beamLength / 2) * x(x > beamLength / 2) - topRightRebar];

% 取正
botLeftRebar = -min(positiveMoment(x <= beamLength / 3));
botMidRebar = -min(positiveMoment(x >= 1 * beamLength / 4 & x <= 3 * beamLength / 4));
botRightRebar = -min(positiveMoment(x >= 2 * beamLength / 3));

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
plot(x, midline, 'Color', grayColor);
hold on;
plot(x, EQ, 'Color', greenColor);
legendEQ = plot(x, NEQ, 'Color', greenColor);
legendGravity = plot(x, DL, 'Color', blueColor);
axis([0 beamLength -50 50]);
legend([legendEQ, legendGravity], 'EQ', 'Gravity', 'Location', 'northeast');
title('Moment');
xlabel('m');
ylabel('tf-m');

% 線性疊加
% 這裡也包含了載重組合的概念
% 取一組載重組合為例
% 1.4DL + E
% 取包絡線
figure;
plot(x, midline, 'Color', grayColor);
hold on;
plot(x, EQ, 'Color', grayColor);
legendEQ = plot(x, NEQ, 'Color', grayColor);
legendGravity = plot(x, DL, 'Color', grayColor);
legendMoment = plot(x, positiveMoment, 'Color', redColor);
plot(x, negativeMoment, 'Color', redColor);
axis([0 beamLength -50 50]);
legend([legendEQ, legendGravity, legendMoment], 'EQ', 'Gravity', 'Linear Add', 'Location', 'northeast');
title('Moment');
xlabel('m');
ylabel('tf-m');

% 依據需求所得出的左中右的配筋
% 左端為 0~1/3 的最大值
% 中央是 1/4~3/4 的最大值
% 右端是 2/3~1 的最大值
% 實際配筋應該會再大一點
% 而且最少會有兩支的限制
% 這裡取最 critical 的情況
figure;
plot(x, midline, '-k');
hold on;
plot(x, EQ, 'Color', grayColor);
plot(x, NEQ, 'Color', grayColor);
plot(x, DL, 'Color', grayColor);
legendMoment = plot(x, positiveMoment, 'Color', redColor);
plot(x, negativeMoment, 'Color', redColor);
legendActural = plot(0, topLeftRebar, 'o', 'Color', redColor);
plot(beamLength / 2, 0, 'o', 'Color', redColor);
plot(beamLength, topRightRebar, 'o', 'Color', redColor);
plot(0, -botLeftRebar, 'o', 'Color', redColor);
plot(beamLength / 2, -botMidRebar, 'o', 'Color', redColor);
plot(beamLength, -botRightRebar, 'o', 'Color', redColor);
axis([0 beamLength -50 50]);
legend([legendActural, legendMoment(1)], 'Actural Rebar', 'Demand', 'Location', 'northeast');
title('Moment');
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
plot(x, EQ, 'Color', grayColor);
plot(x, NEQ, 'Color', grayColor);
plot(x, DL, 'Color', grayColor);
plot(x, positiveMoment, 'Color', grayColor);
plot(x, negativeMoment, 'Color', grayColor);
multiRebar = plot(x, topRebar, 'Color', blueColor);
% plot(x, botRebar, 'Color', blueColor);
legendActural = plot(0, topLeftRebar, 'o', 'Color', redColor);
plot(beamLength / 2, 0, 'o', 'Color', redColor);
plot(beamLength, topRightRebar, 'o', 'Color', redColor);
plot(0, -botLeftRebar, 'o', 'Color', redColor);
plot(beamLength / 2, -botMidRebar, 'o', 'Color', redColor);
plot(beamLength, -botRightRebar, 'o', 'Color', redColor);
axis([0 beamLength -50 50]);
legend([legendActural, multiRebar], 'Actural Rebar', 'Multi Rebar', 'Location', 'northeast');
title('Moment');
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
plot(x, EQ, 'Color', grayColor);
plot(x, NEQ, 'Color', grayColor);
plot(x, DL, 'Color', grayColor);
plot(x, positiveMoment, 'Color', grayColor);
plot(x, negativeMoment, 'Color', grayColor);
multiRebar = plot(x, botRebar, 'Color', blueColor);
legendActural = plot(0, topLeftRebar, 'o', 'Color', redColor);
plot(beamLength / 2, 0, 'o', 'Color', redColor);
plot(beamLength, topRightRebar, 'o', 'Color', redColor);
plot(0, -botLeftRebar, 'o', 'Color', redColor);
plot(beamLength / 2, -botMidRebar, 'o', 'Color', redColor);
plot(beamLength, -botRightRebar, 'o', 'Color', redColor);
axis([0 beamLength -50 50]);
legend([legendActural, multiRebar], 'Actural Rebar', 'Multi Rebar', 'Location', 'northeast');
title('Moment');
xlabel('m');
ylabel('tf-m');

% 而如果我們直接拉直線
% 效益會下降
% 從整體 8.5% 的效益下降到 5.5%
figure;
plot(x, midline, '-k');
hold on;
plot(x, EQ, 'Color', grayColor);
plot(x, NEQ, 'Color', grayColor);
plot(x, DL, 'Color', grayColor);
plot(x, positiveMoment, 'Color', grayColor);
plot(x, negativeMoment, 'Color', grayColor);
multiRebar = plot(x, botRebarOtherMethod, 'Color', blueColor);
legendActural = plot(0, topLeftRebar, 'o', 'Color', redColor);
plot(beamLength / 2, 0, 'o', 'Color', redColor);
plot(beamLength, topRightRebar, 'o', 'Color', redColor);
plot(0, -botLeftRebar, 'o', 'Color', redColor);
plot(beamLength / 2, -botMidRebar, 'o', 'Color', redColor);
plot(beamLength, -botRightRebar, 'o', 'Color', redColor);
axis([0 beamLength -50 50]);
legend([legendActural, multiRebar], 'Actural Rebar', 'Multi Rebar', 'Location', 'northeast');
title('Moment');
xlabel('m');
ylabel('tf-m');

% 所以如果想要從現有配筋就有很好的效果的話 ( 如果已經產生配筋表格了)
% 那我們就會需要更多的資料
figure;
plot(x, midline, '-k');
hold on;
plot(x, EQ, 'Color', grayColor);
legendEQ = plot(x, NEQ, 'Color', grayColor);
legendGravity = plot(x, DL, 'Color', grayColor);
legendMoment = plot(x, positiveMoment, 'Color', grayColor);
plot(x, negativeMoment, 'Color', grayColor);
plot(x, topRebar, 'Color', blueColor);
multiRebar = plot(x, botRebarOtherMethod, 'Color', blueColor);
legendActural = plot(0, topLeftRebar, 'o', 'Color', redColor);
plot(beamLength / 2, 0, 'o', 'Color', redColor);
plot(beamLength, topRightRebar, 'o', 'Color', redColor);
plot(0, -botLeftRebar, 'o', 'Color', redColor);
plot(beamLength / 2, -botMidRebar, 'o', 'Color', redColor);
plot(beamLength, -botRightRebar, 'o', 'Color', redColor);
axis([0 beamLength -50 50]);
legend([legendActural, multiRebar], 'Actural Rebar', 'Multi Rebar', 'Location', 'northeast');
title('Moment');
xlabel('m');
ylabel('tf-m');

% 所以如果想要從現有配筋就有很好的效果的話 ( 如果已經產生配筋表格了)
% 那我們就會需要更多的資料
figure;
plot(x, midline, '-k');
hold on;
plot(x, EQ, 'Color', grayColor);
legendEQ = plot(x, NEQ, 'Color', grayColor);
legendGravity = plot(x, DL, 'Color', redColor);
legendMoment = plot(x, positiveMoment, 'Color', grayColor);
plot(x, negativeMoment, 'Color', grayColor);
legendActural = plot(0, topLeftRebar, 'o', 'Color', redColor);
plot(beamLength / 2, 0, 'o', 'Color', redColor);
plot(beamLength, topRightRebar, 'o', 'Color', redColor);
plot(0, -botLeftRebar, 'o', 'Color', redColor);
plot(beamLength / 2, -botMidRebar, 'o', 'Color', redColor);
plot(beamLength, -botRightRebar, 'o', 'Color', redColor);
axis([0 beamLength -50 50]);
legend([legendActural, legendGravity], 'Actural Rebar', 'Gravity', 'Location', 'northeast');
title('Moment');
xlabel('m');
ylabel('tf-m');

% figure;
% plot(x, midline, '-k');
% hold on;
% legendEQ = plot(x, EQ, '-k', x, NEQ, '-k');
% legendGravity = plot(x, DL, '-k');
% legendMoment = plot(x, negativeMoment, '-r', x, positiveMoment, '-r');
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
% legend([legendEQ(1), legendGravity, legendMoment(1)], 'EQ', 'Gravity', 'Linear Add', 'Location', 'northeast');
% title('Moment');
% xlabel('m');
% ylabel('tf-m');

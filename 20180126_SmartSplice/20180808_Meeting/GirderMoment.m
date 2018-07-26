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

% �t�s�x�ݨD���u
topRebar = [topLeftRebar - topLeftRebar / (beamLength / 2) * x(x <= beamLength / 2), topRightRebar / (beamLength / 2) * x(x > beamLength / 2) - topRightRebar];

% ����
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

% ���O�B�a�_�O����ڻݨD
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

% �u���|�[
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

% ���ݬ� 0~1/3 ���̤j��
% �����O 1/4~3/4 ���̤j��
% �k�ݬO 2/3~1 ���̤j��
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

% ��ڰt�����ӷ|�A�j�@�I
% �ӥB�ַ̤|����䪺����
% �o�̨��� critical �����p
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

% �A�ӴN�O�̾ڲ{���t���̾ڻݨD���u���u��
% �����O�W�h��������
% �����S���ݨD
% ��ݥD�n�ѭ@�_����
% �ڭ̴N�����Ԫ�
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

% ���U�ӬO�U�h��������
% �o�����N��������F
% ���k��ݥѭ@�_����
% �����ڭ̭쥻�w���O�ѭ��O����
% ��ӯu���U�h�����ɭԵo�{�|���a�_�O���]�������i�ӤF
% �p�G�����̷ӭ��O�A��ݨ̾ڦa�_�O���j�ȷ|�p�k���Ŧ⪺�u
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

% �i�H�o�{����骺�����h���F
% ���ⳡ�������O�u
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

% �Ӧp�G�ڭ̪����Ԫ��u
% �įq�|�U��
% �q���� 8.5% ���įq�U���� 5.5%
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

% �ҥH�p�G�Q�n�q�{���t���N���ܦn���ĪG���� ( �p�G�w�g���Ͱt�����F)
% ���ڭ̴N�|�ݭn��h�����
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

% �ҥH�p�G�Q�n�q�{���t���N���ܦn���ĪG���� ( �p�G�w�g���Ͱt�����F)
% ���ڭ̴N�|�ݭn��h�����
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

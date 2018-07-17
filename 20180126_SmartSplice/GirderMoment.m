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

% �t�s�x�ݨD���u
topRebar = [topLeftRebar - topLeftRebar / (beamLength / 2) * x(x <= beamLength / 2), topRightRebar / (beamLength / 2) * x(x > beamLength / 2) - topRightRebar];

% ����
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

% ���O�B�a�_�O����ڻݨD
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

% �u���|�[
% �o�̤]�]�t�F�����զX������
% ���@�ո����զX����
% 1.4DL + E
% ���]���u
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

% �̾ڻݨD�ұo�X�������k���t��
% ���ݬ� 0~1/3 ���̤j��
% �����O 1/4~3/4 ���̤j��
% �k�ݬO 2/3~1 ���̤j��
% ��ڰt�����ӷ|�A�j�@�I
% �ӥB�ַ̤|����䪺����
% �o�̨��� critical �����p
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

% �A�ӴN�O�̾ڲ{���t���̾ڻݨD���u���u��
% �����O�W�h��������
% �����S���ݨD
% ��ݥD�n�ѭ@�_����
% �ڭ̴N�����Ԫ�
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

% ���U�ӬO�U�h��������
% �o�����N��������F
% ���k��ݥѭ@�_����
% �����ڭ̭쥻�w���O�ѭ��O����
% ��ӯu���U�h�����ɭԵo�{�|���a�_�O���]�������i�ӤF
% �p�G�����̷ӭ��O�A��ݨ̾ڦa�_�O���j�ȷ|�p�k���Ŧ⪺�u
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

% �Ӧp�G�ڭ̪����Ԫ��u
% �įq�|�U��
% �q���� 8.5% ���įq�U���� 5.5%
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

% �ҥH�p�G�Q�n�q�{���t���N���ܦn���ĪG���� ( �p�G�w�g���Ͱt�����F)
% ���ڭ̴N�|�ݭn��h�����
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

% �ҥH�p�G�Q�n�q�{���t���N���ܦn���ĪG���� ( �p�G�w�g���Ͱt�����F)
% ���ڭ̴N�|�ݭn��h�����
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

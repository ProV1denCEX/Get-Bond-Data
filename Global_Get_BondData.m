% _Author : Frontal Xiang
%_Version: V 1.0.0
%_Describe:Download Bond Data from website
%_Due to the language of target website, I use Chinese to identify some comments
%***************************************************************
clc
clear all
%% 载入需求
%% Load Config
cNeed = Protected_Load_Needs;

%% 根据不同需求进行数据抓取 主键：Mode
for iNeed = 1 : length(cNeed)
    cThis = cNeed(iNeed);
    
    % 检查是否为新数据下载
    cThis = Protected_Check_IsNew(cThis);
    
    % 获取地址源数据
    cData = Protected_Get_RawData(cThis);
    
    % 重新调整数据结构至需求结构
    cSheet2Write = Protected_Reorganize_Data(cThis, cData);
    
    % 写入Excel -> The person who require this data use Excel to represent some result
	% 		So I write in Excel instead of .mat or csv
    Protected_Write_Excel(cSheet2Write, cThis);

end
disp('******************************* All Data Has Downloaded !!! ************************')
exit;

%% 附属函数区*********************************************************************************
function cThis = Protected_Check_IsNew(cThis)
switch cThis.Mode
    case 'dzzq'
        sExcelName = '单只债券数据.xlsx';
        cThis.IsUpdate = exist(sExcelName, 'file');
        if cThis.IsUpdate
            disp(['Start Updating New ', cThis.Mode, 'Data !'])
            dData = xlsread(sExcelName, 2);
            cThis.DateStart = datenum(num2str(dData(end, 1)), 'yyyymmdd');
        else
            disp(['Start Downloading New ', cThis.Mode, 'Data !'])
        end
        
    case 'zqlb'
        sExcelName = '债券类别';
        nMaturity = cThis.Maturity + 1;
        cMaturity = {'All', 'L01D', 'L02D', 'L03D', 'L07D', 'L14D', 'L21D', 'L01M', 'L02M', 'L03M', 'L04M', 'L06M', 'L09M', 'L01Y'};
        sMaturity = cMaturity{nMaturity};
        sExcelName = [sExcelName, '_', sMaturity, '.xlsx'];
        cThis.IsUpdate = exist(sExcelName, 'file');
        if cThis.IsUpdate
            disp(['Start Updating New ', cThis.Mode, 'Data !'])
            dData = xlsread(sExcelName, 2);
            cThis.DateStart = datenum(num2str(dData(end, 1)), 'yyyymmdd');
        else
            disp(['Start Downloading New ', cThis.Mode, 'Data !'])
        end
        
    case 'tzzlb'
        sExcelName = '投资者类别.xlsx';
        cThis.IsUpdate = exist(sExcelName, 'file');
        if cThis.IsUpdate
            disp(['Start Updating New ', cThis.Mode, 'Data !'])
            dData = xlsread(sExcelName, 2);
            cThis.DateStart = datenum(num2str(dData(end, 1)), 'yyyymmdd');
        else
            disp(['Start Downloading New ', cThis.Mode, 'Data !'])
        end
        
    otherwise
        
end
end


function Protected_Write_Excel(cSheet2Write, cThis)
switch cThis.Mode
    case 'dzzq'
        sExcelName = '单只债券数据';
        for iSheet = 1 : size(cSheet2Write, 1)
            sSheetName = cSheet2Write{iSheet, 1};
            cContent = cSheet2Write{iSheet, 2};
            if cThis.IsUpdate
                [~, ~, cTemp] = xlsread(sExcelName, sSheetName);
                [~, dLocation_Out, dLocation_In] = intersect(cTemp(1, :), cContent(1, :));
                cNew = cell(size(cTemp, 1), size(cContent, 2));
                [cNew{:}] = deal(0);
                cNew(2 : end, dLocation_In) = cTemp(2 : end, dLocation_Out);
                cNew(1, :) = cContent(1, :);
                cContent = [cNew(1 : end - 1, :); cContent(2 : end, :)];
            else
            end
            xlswrite([sExcelName, '.xlsx'], cContent, sSheetName);
        end
        disp([cThis.Mode, 'Data Finished !'])
        
    case 'zqlb'
        sExcelName = '债券类别';
        nMaturity = cThis.Maturity + 1;
        cMaturity = {'All', 'L01D', 'L02D', 'L03D', 'L07D', 'L14D', 'L21D', 'L01M', 'L02M', 'L03M', 'L04M', 'L06M', 'L09M', 'L01Y'};
        sMaturity = cMaturity{nMaturity};
        sExcelName = [sExcelName, '_', sMaturity, '.xlsx'];
        for iSheet = 1 : length(cSheet2Write)
            sSheetName = cSheet2Write{iSheet, 1};
            cContent = cSheet2Write{iSheet, 6};
            if cThis.IsUpdate
                [~, ~, cTemp] = xlsread(sExcelName, sSheetName);
                cContent = [cTemp(1 : end - 1, :); cContent(2 : end, :)];
            else
            end
            xlswrite(sExcelName, cContent, sSheetName);
        end
        disp([cThis.Mode, 'Data Finished !'])
        
    case 'tzzlb'
        sExcelName = '投资者类别';
        for iSheet = 1 : length(cSheet2Write)
            sSheetName = cSheet2Write{iSheet, 1};
            cContent = cSheet2Write{iSheet, 5};
            if cThis.IsUpdate
                [~, ~, cTemp] = xlsread(sExcelName, sSheetName);
                cContent = [cTemp(1 : end - 1, :); cContent(2 : end, :)];
            else
            end
            xlswrite([sExcelName, '.xlsx'], cContent, sSheetName);
        end
        disp([cThis.Mode, 'Data Finished !'])
        
    otherwise
        
end
end


function cSheet2Write = Protected_Reorganize_Data(cThis, cData)
cData2Write = cData;
dLocated = cellfun(@(x) ~isempty(x), cData2Write(:, 2));
cData2Write = cData2Write(dLocated, :);

switch cThis.Mode
    case 'dzzq'
        nTemp = sum(cellfun(@(x) size(x, 1), cData2Write(:, 2)));
        cCode = cell(nTemp, 1);
        cName = cell(nTemp, 1);
        nLocation_Start = 1;
        for iDate = 1 : size(cData2Write, 1)
            cTemp = cData2Write{iDate, 2};
            cCode(nLocation_Start : nLocation_Start + size(cTemp, 1) - 1) = cTemp(:, 2);
            cName(nLocation_Start : nLocation_Start + size(cTemp, 1) - 1) = cTemp(:, 1);
            nLocation_Start = nLocation_Start + size(cTemp, 1);
        end
        cName = [cName, cCode];
        cCode = unique(cCode);
        [~, dLocation] = sort(cellfun(@(x) str2double(x), cCode));
        cCode = cCode(dLocation);
        
        dTimeAxis = str2num(datestr(cell2mat(cData2Write(:, 1)), 'yyyymmdd'));
        dData2Write = zeros(length(dTimeAxis), size(cCode, 1) * size(cData2Write{1, 4}, 2));
        for iDate = 1 : length(dTimeAxis)
            cTemp = cData2Write{iDate, 2};
            for iCode = 1 : size(cTemp, 1)
                sCode = cTemp{iCode, 2};
                dLocation = find(strcmp(sCode, cCode));
                dData2Write(iDate, dLocation) = str2double(cTemp{iCode, 3});
                dData2Write(iDate, dLocation + length(cCode)) = str2double(cTemp{iCode, 4});
            end
        end
        dData2Write = [dTimeAxis, dData2Write];
        cSheet2Write = cData2Write{1, 4}';
        cSheet2Write{1, 2} = dData2Write(:, [1, 2 : length(cCode) + 1]);
        cSheet2Write{2, 2} = dData2Write(:, [1, length(cCode) + 2 : end]);
        
        cName_All = cell(1, length(cCode));
        for iCode = 1 : length(cCode)
            sCode = cCode{iCode};
            dLocation = find(strcmp(cName(:, 2), sCode), 1, 'first');
            cName_All{iCode} = [cName{dLocation, 2}, '_(', cName{dLocation, 1}, ')'];
        end
        cIndex = ['日期', cName_All];
        cSheet2Write{1, 2} = [cIndex; num2cell(cSheet2Write{1, 2})];
        cSheet2Write{2, 2} = [cIndex; num2cell(cSheet2Write{2, 2})];
        
    case 'zqlb'
        dTimeAxis = str2num(datestr(cell2mat(cData2Write(:, 1)), 'yyyymmdd'));
        dData2Write = zeros(length(dTimeAxis), size(cData2Write{1, 2}, 1) * size(cData2Write{1, 2}, 2));
        for iDate = 1 : length(dTimeAxis)
            cTemp = cData2Write{iDate, 2};
            dLocated = cellfun(@(x) ischar(x), cTemp);
            cTemp(dLocated) = num2cell(cellfun(@(x) str2double(x), cTemp(dLocated)));
            dData = cell2mat(cTemp);
            dData = reshape(dData(1 : size(cData2Write{1, 2}, 1), 1 : size(cData2Write{1, 2}, 2)), 1, size(dData2Write, 2));
            dData2Write(iDate, :) = dData;
        end
        dData2Write = [dTimeAxis, dData2Write];
        
        cSheet2Write = cData2Write{1, 3};
        for iSheet = 1 : length(cSheet2Write)
            for iIndex = 1 : size(cData2Write{1, 2}, 2)
                cSheet2Write{iSheet, iIndex + 1} = dData2Write(:, [1, iSheet + 1 + length(cSheet2Write) * (iIndex - 1)]);
            end
            dTemp = [cSheet2Write{iSheet, 2 : 5}];
            dTemp = dTemp(:, [1, 2 * (1 : size(cData2Write{1, 2}, 2))]);
            cTemp = num2cell(dTemp);
            cIndex = ['日期', cData2Write{1, 4}];
            cSheet2Write{iSheet, 6} = [cIndex; cTemp];
        end
        
    case 'tzzlb'
        dTimeAxis = str2num(datestr(cell2mat(cData2Write(:, 1)), 'yyyymmdd'));
        dData2Write = zeros(length(dTimeAxis), size(cData2Write{1, 2}, 1) * size(cData2Write{1, 2}, 2));
        for iDate = 1 : length(dTimeAxis)
            cTemp = cData2Write{iDate, 2};
            dLocated = cellfun(@(x) ischar(x), cTemp);
            cTemp(dLocated) = num2cell(cellfun(@(x) str2double(x), cTemp(dLocated)));
            dData = cell2mat(cTemp);
            dData = reshape(dData, 1, size(dData2Write, 2));
            dData2Write(iDate, :) = dData;
        end
        dData2Write = [dTimeAxis, dData2Write];
        
        cSheet2Write = cData2Write{1, 3};
        for iSheet = 1 : length(cSheet2Write)
            for iIndex = 1 : size(cData2Write{1, 2}, 2)
                cSheet2Write{iSheet, iIndex + 1} = dData2Write(:, [1, iSheet + 1 + length(cSheet2Write) * (iIndex - 1)]);
            end
                 cSheet2Write{iSheet, 4} = cData2Write{1, 4};
                 dTemp = [cSheet2Write{iSheet, 2}, cSheet2Write{iSheet, 3}];
                 dTemp = dTemp(:, [1, 2 * (1 : size(cData2Write{1, 2}, 2))]);
                 cTemp = num2cell(dTemp);
                 cIndex = ['日期', cSheet2Write{iSheet, 4}];
                 cSheet2Write{iSheet, 5} = [cIndex; cTemp];
        end
        
    otherwise

end
end


function cData_all = Protected_Get_RawData(cThis)
cData_all = num2cell(cThis.DateStart : cThis.DateEnd)';
[cData_all{:, 2 : 4}] = deal([]);

nFrequency = cThis.Frequency;
nIssuer = cThis.Issuer;
nInterestType = cThis.InterestType;
nMaturity = cThis.Maturity;
nCode = cThis.Code;
sMode = cThis.Mode;

cParpool = parpool;
parfor iDate = 1 : size(cData_all, 1)
    cTemp = cData_all(iDate, :);
    dDate = datevec(cTemp{1});
    nYear = dDate(1);
    nMonth = dDate(2);
    nDay = dDate(3);

    sContent = Fun_Get_Data(nFrequency, ...
        nYear, ...
        nMonth, ...
        nDay, ...
        nIssuer, ...
        nInterestType, ...
        nMaturity, ...
        nCode, ...
        sMode);
    
    [cData, cSheetList, cIndexName] = Fun_Screen_Data(sContent, sMode);
    cTemp{2} = cData;
    cTemp{3} = cSheetList;
    cTemp{4} = cIndexName;
    
    cData_all(iDate, :) = cTemp;
    disp([datestr(cTemp{1}, 'yyyymmdd'), ' ''s ', sMode, ' Data Has Downloaded !'])
end
delete(cParpool);
end


function [cData, cSheetList, cIndexName] = Fun_Screen_Data(sContent, sMode)
switch sMode
    case 'dzzq'
        % IsTradeDay & SheetName
        sExpr = ['<td align="center">', '.*?', '</td>'];
        cTemp = regexp(sContent, sExpr, 'match');
        cTemp_SheetName = cellfun(@(x) x(20 : end - 5), cTemp, 'UniformOutput', 0);
        if ~isempty(cTemp_SheetName)
            cIndexName = cTemp_SheetName(3 : 4);
        else
            cData = [];
            cSheetList = 0;
            cIndexName = [];
            return
        end

        % Data
        sExpr = ['<td align = right  nowrap>', '.*?', '</td>'];
        cTemp = regexp(sContent, sExpr, 'match');
        cTemp = cellfun(@(x) x(27 : end - 5), cTemp, 'UniformOutput', 0);
        cTemp = reshape(cTemp, 5, length(cTemp) / 5);
        cData = cTemp';
        cData = cData(2 : end, :);
        cSheetList = 1;
        
    case 'zqlb'
        % Bond Type
        sExpr = ['<td align = left   nowrap>', '.*?', '</td>'];
        cTemp = regexp(sContent, sExpr, 'match');
        cTemp_Type = cellfun(@(x) x(27 : end - 5), cTemp, 'UniformOutput', 0);
        if ~isempty(cTemp_Type)
        else
            cData = [];
            cSheetList = [];
            cIndexName = [];
            return
        end
        
        % SheetName
        sExpr = ['<td align="center">', '.*?', '</td>'];
        cTemp = regexp(sContent, sExpr, 'match');
        cTemp_SheetName = cellfun(@(x) x(20 : end - 5), cTemp, 'UniformOutput', 0);
        if ~isempty(cTemp_SheetName)
            cIndexName = cTemp_SheetName(2 : end);
        else
            cData = [];
            cIndexName = [];
            return
        end
        
        % Data
        sExpr = ['<td align = right  nowrap>', '.*?', '</td>'];
        cTemp = regexp(sContent, sExpr, 'match');
        cTemp_Data = cellfun(@(x) x(27 : end - 5), cTemp, 'UniformOutput', 0);
        if ~isempty(cTemp_Data)
        else
            cData = [];
            return
        end
        
        % Insert
        cTemp = reshape(cTemp_Data, length(cTemp_Data) / length(cTemp_Type), length(cTemp_Type));
        cData = cTemp';
        cData = cData(2 : end, :);
        dLocation = cellfun(@(x) isempty(x), cData);
        [cData{dLocation}] = deal(0);
        cSheetList = cTemp_Type';
        cSheetList = cSheetList(2 : end, :);
        
    case 'tzzlb'
        % Investor Type
        sExpr = ['<td align = left   nowrap>', '.*?', '</td>'];
        cTemp = regexp(sContent, sExpr, 'match');
        cTemp_Type = cellfun(@(x) x(27 : end - 5), cTemp, 'UniformOutput', 0);
        if ~isempty(cTemp_Type)
        else
            cData = [];
            cSheetList = [];
            cIndexName = [];
            return
        end
        
        % SheetName
        sExpr = ['<td align="center">', '.*?', '</td>'];
        cTemp = regexp(sContent, sExpr, 'match');
        cTemp_SheetName = cellfun(@(x) x(20 : end - 5), cTemp, 'UniformOutput', 0);
        if ~isempty(cTemp_SheetName)
            cIndexName = cTemp_SheetName(2 : end);
        else
            cData = [];
            cIndexName = [];
            return
        end

        % Data
        sExpr = ['<td align = right  nowrap>', '.*?', '</td>'];
        cTemp = regexp(sContent, sExpr, 'match');
        cTemp_Data = cellfun(@(x) x(27 : end - 5), cTemp, 'UniformOutput', 0);
        if ~isempty(cTemp_Data)
        else
            cData = [];
            return
        end
        
        % Insert
        cTemp = reshape(cTemp_Data, length(cTemp_Data) / length(cTemp_Type), length(cTemp_Type));
        cData = cTemp';
        cData = cData(2 : end, :);
        dLocation = cellfun(@(x) isempty(x), cData);
        [cData{dLocation}] = deal(0);
        cSheetList = cTemp_Type';
        cSheetList = cSheetList(2 : end, :);

    otherwise
        
end
end

% _Author : Frontal Xiang
%_Version: V 1.0.0
%_Describe: 根据需求下载债券借贷数据
%_Update: 20180112 完成代码
%_Input: 
%       nFrequency  时间频率 1日 2周 3月 4季 5年
%       nYear 选择的时间 年
%       nMonth 选择的时间 月
%       nDay 选择的时间 日
%       nIssuer 发行人 0-8 全部 财政 国开 进出口行 人总行 中信信托 其他 农发 铁道部
%       nInterestType 计息方式 99无计息 10贴现 20利随本金 31固定 32浮动 40无息
%       nMaturity 借贷期限 -》sMaturity
%       nCode 债券代码 无则为空
%       sMode dzzq 单只债券 zqlb 债券类别
%_Output:null
%*********************************************************************************
function sContent = Fun_Get_Data(nFrequency, nYear, nMonth, nDay, nIssuer, nInterestType, nMaturity, nCode, sMode)
sContent = [];
nMaturity = nMaturity + 1;
cMaturity = {'00', 'L01D', 'L02D', 'L03D', 'L07D', 'L14D', 'L21D', 'L01M', 'L02M', 'L03M', 'L04M', 'L06M', 'L09M', 'L01Y'};

sMaturity = cMaturity{nMaturity};

sHeader = 'Mozilla/5.0 (Windows NT 10.0; WOW64; rv:60.0) Gecko/20100101 Firefox/60.0';
sUrl = ['http://www.chinabond.com.cn/jsp/include/EJB/jdtj_', sMode, ...
    '.jsp?', ...
    'sel4=', num2str(nFrequency), ...                           % 时间频率 1日 2周 3月 4季 5年
    '&tbSelYear6=', num2str(nYear), ...         % 选择的时间 年
    '&tbSelMonth6=', num2str(nMonth), ...           % 选择的时间 月
    '&calSelectedDate6=', num2str(nDay), ...      % 选择的时间 日
    '&ZQFXRJD1=', num2str(nIssuer, '%02d'), ...             % 发行人
    '&FUXFSJD1=', num2str(nInterestType, '%02d'), ...             % 付息方式
    '&JXFSJD2=', num2str(nInterestType, '%02d'), ...                % 计息方式
    '&JDQX2=', sMaturity, ...                    % 借贷期限
    '&ZQFXRJD3=', num2str(nIssuer, '%02d'), ...             %  投资人类别
    '&ZQFXRJD4=', num2str(nIssuer, '%02d'), ...             % 单只债券 发行人
    '&I_ZQDM_JD=', num2str(nCode)];                 % 单只债券 代码

nTimes = 0;
while true
    [sContent, nStatus] = urlread(sUrl, 'Timeout', 30, 'UserAgent', sHeader);
    
    
    if nStatus || nTimes >= 10
        break
    else
        nTimes = nTimes + 1;
    end
end
end

function cNeed = Protected_Load_Needs
% Date - Check if Update
nDateStart = 20100101;
nDateEnd = str2double(datestr(now, 'yyyymmdd'));
nDateStart = datenum(num2str(nDateStart), 'yyyymmdd');
nDateEnd = datenum(num2str(nDateEnd), 'yyyymmdd');

% dzzq
cNeed.Mode = 'dzzq';
cNeed.Frequency = 1;
cNeed.Issuer = 0;
cNeed.InterestType = 0;
cNeed.Maturity = 0;
cNeed.Code = [];
cNeed.DateStart = nDateStart;
cNeed.DateEnd = nDateEnd;
cNeed.IsUpdate = 1;

% zqlb
dMaturity = 0 : 13;
for iNeed = 1 : length(dMaturity)
    cNeed(end + 1).Mode = 'zqlb';
    cNeed(end).Frequency = 1;
    cNeed(end).Issuer = 0;
    cNeed(end).InterestType = 0;
    cNeed(end).Maturity = dMaturity(iNeed);
    cNeed(end).Code = [];
    cNeed(end).DateStart = nDateStart;
    cNeed(end).DateEnd = nDateEnd;
    cNeed(end).IsUpdate = 1;
end

% tzzlb
cNeed(end + 1).Mode = 'tzzlb';
cNeed(end).Frequency = 1;
cNeed(end).Issuer = 0;
cNeed(end).InterestType = 0;
cNeed(end).Maturity = 0;
cNeed(end).Code = [];
cNeed(end).DateStart = nDateStart;
cNeed(end).DateEnd = nDateEnd;
cNeed(end).IsUpdate = 1;
end
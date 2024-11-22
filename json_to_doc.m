fclose('all');
%%
clear all; clc; close all;
%%
path = '';
path_save = '';
if ~exist(path_save, 'dir')
    mkdir(path_save)
end

jsonFileName = '.json'; %filename
jsonText = fileread([path,jsonFileName]);
data = jsondecode(jsonText);


% Create Word application object
Word = actxserver('Word.Application');
Word.Visible = true;

% Create a new document
doc = Word.Documents.Add;

% Create a table
numRows = length(data.messages) + 1; % +1 for header row
numCols = 4;
range = doc.Content;
table = doc.Tables.Add(range, numRows, numCols);

% Set table headers
headers = {'Timestamp', 'Sender Name', 'Text', 'Media'};
for i = 1:length(headers)
    table.Cell(1, i).Range.Text = headers{i};
end

% Populate table with data
for i = 1:numRows
    time = datetime(data.messages(i).timestamp/1000, 'ConvertFrom', 'posixtime');
    table.Cell(i+1, 1).Range.Text = string(time);
    table.Cell(i+1, 2).Range.Text = data.messages(i).senderName;
    table.Cell(i+1, 3).Range.Text = data.messages(i).text;
    
    if strcmp(data.messages(i).type, 'media')
        table.Cell(i+1, 3).Range.Text = data.messages(i).type;
        table.Cell(i+1, 4).Range.Text = data.messages(i).media.uri;
    else
        table.Cell(i+1, 4).Range.Text = '';
    end
end

% Save the document
doc.SaveAs2([path_save 'chat_log.docx']);

%% Close Word
Word.Quit;

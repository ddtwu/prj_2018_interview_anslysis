#---------------------------------------------------------#
source('E:/Marvin.Wu/myScript/initial_func.R', encoding = 'UTF-8')
library(jiebaR)
library(wordcloud2)




#---------------------------------------------------------#
#---------------------------------------------------------#
#---------------------------------------------------------#
# 讀取原始資料 from 文本清單
docList <- read_xlsx('E:/Marvin.Wu/myProject/2018_BAC/2018_訪談分析/doc list.xlsx') %>% distinct()

# > 確認訪談篇數
n_doc <- docList %>% dim() %>% .[1]



#---------------------------------------------------------#
#---------------------------------------------------------#
#---------------------------------------------------------#
# 資料清整程序
rawdata_all <- NULL
for (i in 1:n_doc) {
  
  Text_ID <- docList %>% slice(i) %>% pull(Text_ID)
  Text_Data <- readLines(sprintf('E:/Marvin.Wu/myProject/2018_BAC/2018_訪談分析/%s.txt', Text_ID))
  
  #---------------------------#
  # 資料讀取
  rawdata <- Text_Data[Text_Data != ""] %>% data.frame(textContent = ., stringsAsFactors = FALSE) %>% tbl_df()
  
  # > 填上文本ID以及句子的原始排序
  rawdata %<>% mutate(Text_ID = Text_ID, sentSeq = row_number())
  
  # > 將原始排序轉換為前補0的字串
  tmp <- rawdata %>% pull(sentSeq) %>% max() %>% str_length()
  rawdata %<>% 
    mutate(sentSeq = str_pad(sentSeq, pad = '0', width = tmp, side = 'left'))
  rm(tmp)
  
  # > 轉換全形字
  rawdata %<>% mutate(textContent = str_replace_all(textContent, "Ｑ", "Q"),
                      textContent = str_replace_all(textContent, "Ａ", "A"))
  
  # > 區分該句是提問或是回答
  rawdata %<>% mutate(sentType = ifelse(str_detect(textContent, "#Q："), 'q', 'a'))
  
  # > 區分段落
  tmp <- rawdata %>% pull(sentType)
  qq <- str_which(tmp, 'q')
  
  bb <- str_which(tmp, 'q') %>% lead()
  bb <- bb[!is.na(bb)] %>% c(., length(tmp)+1)
  
  paraTitle <- str_c('para', 1:length(qq)) %>% rep(x = ., times = (bb-qq))
  rawdata %<>% mutate(paraGrp = paraTitle)
  
  # > 建立句子的ID
  rawdata %<>% group_by(sentType, paraGrp) %>% mutate(seqTmp = row_number()) %>% ungroup() 
  # rawdata %<>% mutate(tmp = str_sub(sentType, 1, 1)) %>% mutate(seq_tmp = str_c(tmp, seqTmp))
  rawdata %<>% mutate(sentId = str_c(Text_ID, sentSeq, paraGrp, sentType, seqTmp, sep = '_')) %>% select(-seqTmp, -sentSeq)
  
  # > 表格最後的整理
  rawdata %<>% 
    mutate(textContent = str_replace_all(textContent, "Q：", "")) %>% 
    mutate(textContent = str_replace_all(textContent, "#", "")) %>% 
    mutate(textContent = str_replace_all(textContent, "A：", "")) %>% 
    mutate(textContent = str_replace_all(textContent, " ", "")) %>% 
    mutate(textContent = str_to_lower(textContent))
  
  rawdata %<>% select(Text_ID, sentId, paraGrp, sentType, textContent)
  
  # > 合併
  rawdata_all %<>% bind_rows(rawdata)
}



#---------------------------------------------------------#
#---------------------------------------------------------#
#---------------------------------------------------------#
# 斷詞斷字
# set jieba worker
jworker <- worker()

# > 增加特定的詞
newWord <- c('')
new_user_word(jworker, newWord)

# > 只處理回答的部分
sub_data <- rawdata_all %>% filter(sentType == 'a')

# > 逐句進行斷詞斷字
Text_ID <- sub_data$Text_ID
sentId <- sub_data$sentId
rstTmp <- apply_list(input = as.list(sub_data$textContent), worker = jworker)

# > 計算詞頻
freqTb <- NULL
for (j in 1:length(rstTmp)) {
  
  freqTmp <- rstTmp[[j]] %>% 
    freq() %>% 
    tbl_df() %>% 
    mutate(Text_ID = Text_ID[j], sentId = sentId[j], lengChar = str_length(char)) 
  freqTb %<>% bind_rows(freqTmp)
}
freqTb %<>% select(Text_ID, sentId, char, freq, lengChar)



#---------------------------------------------------------#
#---------------------------------------------------------#
#---------------------------------------------------------#
# 詞頻統計表
# > 整體
rstTb1 <- freqTb %>%
  group_by(char, lengChar) %>% 
  summarise(sumFreq = sum(freq)) %>% 
  ungroup() %>% 
  # 排除單詞
  filter(lengChar > 1) %>% 
  arrange(desc(sumFreq)) %>% 
  mutate(seq_freq = row_number())

# > 區分Text_ID
rstTb2 <- freqTb %>%
  group_by(Text_ID, char, lengChar) %>% 
  summarise(sumFreq = sum(freq)) %>% 
  ungroup() %>% 
  # 排除單詞
  filter(lengChar > 1) %>% 
  arrange(Text_ID, desc(sumFreq)) %>% 
  group_by(Text_ID) %>% 
  mutate(seq_freq = row_number()) %>% 
  ungroup()



#---------------------------------------------------------#
#---------------------------------------------------------#
#---------------------------------------------------------#
# 對應特定關鍵字的句子查詢
tmp <- freqTb %>% filter(char == '數位') %>% select(Text_ID, sentId) %>% distinct() # %>% unite(index, Text_ID, sentId)

# > 有出現特定關鍵字的句子：原始內容
rawdata_all %>%
  # unite(index, Text_ID, sentId, remove = FALSE) %>% 
  filter(sentId %in% tmp$sentId) %>% View
  # pull(textContent)

# > 有出現特定關鍵字的句子：詞頻彙整
freqTb %>% 
  filter(lengChar > 1) %>%
  # unite(index, Text_ID, sentId, remove = FALSE) %>% 
  filter(sentId %in% tmp$sentId) %>% 
  # select(-index) %>% 
  group_by(Text_ID, char, lengChar) %>% 
  summarise(sumFreq = sum(freq)) %>% 
  ungroup() %>% 
  arrange(desc(sumFreq))



#---------------------------------------------------------#
#---------------------------------------------------------#
#---------------------------------------------------------#
# 文字雲
library(wordcloud2)

# > 先濾掉1個字的 & 濾掉頻率太低的詞(<= 1)
### 審核過的願望
rstTb1 %>% filter(lengChar > 1 & sumFreq > 1) %>% select(char, sumFreq) %>% data.frame() %>% 
  wordcloud2(color = 'random-dark', rotateRatio = 0, size = 1.5, gridSize = 1, shuffle = FALSE)



#---------------------------------------------------------#
#---------------------------------------------------------#
#---------------------------------------------------------#
# export
sheetsList <- list('訪談清單' = data.frame(docList),
                   '訪談紀錄彙整' = data.frame(rawdata_all),
                   '斷詞結果' = data.frame(freqTb),
                   '詞頻統計(排除單詞)' = data.frame(rstTb2))

write_xlsx(sheetsList,
           "E:/Marvin.Wu/myProject/2018_BAC/2018_訪談分析/test_result_20181221.xlsx")







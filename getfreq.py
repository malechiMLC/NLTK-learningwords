import nltk
import string
from nltk.stem.wordnet import WordNetLemmatizer     # 词形还原
from nltk.tokenize import word_tokenize     # 获取单词的TAG
from nltk.corpus import wordnet     # 通过tag获取pos
from nltk import FreqDist       # 统计词频
import xlwt     # 写入excel

# 运行一次即可
# nltk.download('punkt')
# nltk.download('averaged_perceptron_tagger')
# nltk.download('wordnet')


# 打开文件对象
# 为了避免忘记close掉这个文件对象


def open_txt(srcFile):
    with open(srcFile, "r", encoding='UTF-8') as f:
        sentence = f.read()
    return sentence

# 去掉/替换sentence中不想要的字符
# sentence是要被替换的字符串
# 用outtab中对应的每个值替换intab中对应的值
# 一对一替换，长度需要一样
# deletetab是要删除的值，包含的每个值都将被删除
# 去除标点是，deletetab为string.punctuation


def translate(sentence, intab, outtab, deletetab):
    trantab = str.maketrans(intab, outtab, deletetab)
    return sentence.translate(trantab).lower()


# 分词sentence->tokens
# tokens是一个元素类型为str的list


def tokenize_words(sentence):
    return nltk.word_tokenize(sentence)


# 获取tokens中每个词的词形tag


def get_pos_tag(tokens):
    return nltk.pos_tag(tokens)

# 根据tag获取pos
# 作为后面词形还原函数的参数


def get_wordnet_pos(treebank_tag):
    if treebank_tag.startswith('J'):    # 形容词
        return wordnet.ADJ
    elif treebank_tag.startswith('V'):  # 动词
        return wordnet.VERB
    elif treebank_tag.startswith('N'):  # 名词
        return wordnet.NOUN
    elif treebank_tag.startswith('R'):  # 副词
        return wordnet.ADV
    elif treebank_tag.startswith('C'):  # 连词，数词：跳过
        return 'delete'
    elif treebank_tag.startswith('D'):  # 限定词：跳过
        return 'delete'
    elif treebank_tag.startswith('D'):  # 存在量词：跳过
        return 'delete'
    elif treebank_tag.startswith('I'):  # 介词，连接词：跳过
        return 'delete'
    elif treebank_tag.startswith('L'):  # 标记：跳过
        return 'delete'
    elif treebank_tag.startswith('M'):  # 情态动词：跳过
        return 'delete'
    elif treebank_tag.startswith('P'):  # 前限定词，所有格标记，人称代词，所有格：跳过
        return 'delete'
    elif treebank_tag.startswith('T'):  # to：跳过
        return 'delete'
    elif treebank_tag.startswith('W'):  # Wh限定词，WH代词，WH代词所有格，WH副词：跳过
        return 'delete'
    else:
        return 'else'

# 词形还原


def lemmatize_words(tagged_tokens):
    lemmatized_tokens = []
    lemmatizer = WordNetLemmatizer()
    for tagged_token in tagged_tokens:
        if get_wordnet_pos(tagged_token[1]) == 'delete':
            continue
        elif get_wordnet_pos(tagged_token[1]) == 'else':
            lemmatized_tokens.append(tagged_token[0])
        else:
            lemmatized_tokens.append(lemmatizer.lemmatize(
                tagged_token[0], get_wordnet_pos(tagged_token[1])))
    return lemmatized_tokens


# 去除过短的词


def filt_short_words(tokens):
    i = 0
    while i < len(tokens):
        if len(tokens[i]) <= 3:
            tokens.pop(i)
            i -= 1
        i += 1
    return tokens


# 统计词频
# 按词频由高到低排列


def sort_freq(lemmatized_tokens):
    fdist = FreqDist(lemmatized_tokens)     # 统计词频
    length = len(list(fdist))       # 获取总词数
    freq_list = fdist.most_common(length)       # 按词频由高到低排序
    return freq_list

# 写入excel


def write_to_excel(excelPath, freq_list):
    workbook = xlwt.Workbook(encoding='utf-8')      # 创建一个workbook 设置编码
    worksheet = workbook.add_sheet('sheet')     # 创建一个worksheet
    row = 0
    length = len(list(freq_list))
    while row <= length:        # 参数对应 行, 列, 值
        worksheet.write(row, 0, label=str(freq_list[row-1][1]))
        worksheet.write(row, 1, label=str(freq_list[row-1][0]))
        row += 1
    workbook.save(excelPath)    # 保存，会写入当前项目目录下，没有的话会新建一个文件


def get_freq_xl(srcFile, intab, outtab, deletetab, excelPath):
    sentence = open_txt(srcFile)
    sentence = translate(sentence, intab, outtab, deletetab)
    tokens = tokenize_words(sentence)
    tagged_tokens = get_pos_tag(tokens)
    lemmatized_tokens = lemmatize_words(tagged_tokens)
    lemmatized_tokens = filt_short_words(lemmatized_tokens)
    freq_list = sort_freq(lemmatized_tokens)
    write_to_excel(excelPath, freq_list)


# srcFile = 'C:/Users/30544/Desktop/test.txt'

# sentence = open_txt(srcFile)

# sentence = translate(sentence, "", "", string.punctuation)

# tokens = tokenize_words(sentence)

# tagged_tokens = get_pos_tag(tokens)

# lemmatized_tokens = lemmatize_words(tagged_tokens)

# lemmatized_tokens = filt_short_words(lemmatized_tokens)

# freq_list = sort_freq(lemmatized_tokens)

# write_to_excel('vocabularies.xls', freq_list)

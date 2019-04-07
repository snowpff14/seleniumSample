from webBase import WebExecuteBase
import sys, os
import pandas as pd
import numpy as np
import unittest
import pathlib
from seleniumTestSite2 import SeleniumTestSite




class SeleniumTestSiteOrderTestSuite(unittest.TestCase):

    # 前処理クエリクローズのテスト時はクローズ対象のクエリを開く
    def setUp(self):
        print('準備処理開始')
        print('準備処理終了')


    # 後処理クエリ発行のテスト時は発行したクエリを取り消す
    def tearDown(self):
        print('後処理開始')
        print('後処理終了')

    # サンプルとして用意してあるファイルの内容との一致を確認
    # ベースとなるファイル
    def testSeleniumTestSiteCompare(self):
        # テスト用として実行する
        resourceFiles=[ 'unitTest/resources/001/appConfig.ini']
        seleniumTestSite =SeleniumTestSite()
        seleniumTestSite.init('unitTest/data/001/予約データ2.xlsx',mode=1,filePaths=resourceFiles)
        seleniumTestSite.mainExecute()
        resultFilePath=seleniumTestSite.getResultPath()

        # 比較用の結果はCSVが1番目、Excelが2番目に格納されている
        actualCSVFile=pathlib.Path(resultFilePath[0])
        actualExcelFile=pathlib.Path(resultFilePath[1])

        # 想定された結果が取得できた時
        resultFilePathCSVExpect=pathlib.Path('unitTest/result/001/宿泊情報_テスト結果.csv')
        resultFilePathExcelExpect=pathlib.Path('unitTest/result/001/宿泊情報_テスト結果.xlsx')
        # CSV ファイルを読み込んで比較を行う
        self.csvFileCompare(actualCSVFile,resultFilePathCSVExpect)
        self.excelFileCompare(actualExcelFile,resultFilePathExcelExpect,'Sheet1')


        
    # 登録にすべて失敗したときのパターン
    def testEmptyCompare(self):
        resourceFiles=[ 'unitTest/resources/002/appConfig.ini']
        seleniumTestSite =SeleniumTestSite()
        seleniumTestSite.init('unitTest/data/002/予約データEmpty.xlsx',mode=1,filePaths=resourceFiles)
        seleniumTestSite.mainExecute()
        resultFilePath=seleniumTestSite.getResultPath()
        self.assertEqual(len(resultFilePath),0)



    ## 別途作成したユニットテスト用の比較処理をこちらに移植
    # CSVファイルの比較用
    def csvFileCompare(self,actualFilePath,expectFilePath):
        actualDataSets = pd.read_csv(actualFilePath,dtype='str',encoding='cp932')
        expectDataSets = pd.read_csv(expectFilePath,dtype='str',encoding='cp932')

        actualDataArray=np.asarray(actualDataSets)
        expectDataArray=np.asarray(expectDataSets)
        self.bothNumpyArrayEqual(actualDataArray,expectDataArray)

    # エクセルファイルの比較用
    def excelFileCompare(self,actualFilePath,expectFilePath,sheetName):
        actualDataSets = pd.read_excel(actualFilePath,dtype='str',sheet_name=sheetName)
        expectDataSets = pd.read_excel(expectFilePath,dtype='str',sheet_name=sheetName)

        actualDataArray=np.asarray(actualDataSets)
        expectDataArray=np.asarray(expectDataSets)
        self.bothNumpyArrayEqual(actualDataArray,expectDataArray)


    # pandasのDataFrameをnumpyのArrayに変換したものを比較する
    def bothNumpyArrayEqual(self,actualData,expectData):
        for index,expect in enumerate(expectData):
            self.assertEqual(expect.tolist(),actualData[index].tolist())



if __name__ == '__main__':
    unittest.main()
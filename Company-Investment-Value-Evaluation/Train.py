from openpyxl import load_workbook
import pandas as pd
import csv
import seaborn as sns
import numpy as np
import lightgbm as lgb
from sklearn.decomposition import PCA
from sklearn.model_selection import KFold
from sklearn import preprocessing
from sklearn import svm
from sklearn.ensemble import RandomForestRegressor
import matplotlib.pyplot as plt

CORR_LIMIT = 0.7


# Lightgbm-GBDT
def GBDT(X_data, X_label, y_data, num_leaves=23, min_data_in_leaf=8, feature_fraction=0.7, bagging_fraction=0.5):
    features = [c for c in X_data.columns if c not in ['企业编号']]

    X_data = X_data.values

    params = {
        'task': 'train',
        'boosting_type': 'gbdt',
        'objective': 'regression',
        'metric': {'l2_root'},
        'num_leaves': num_leaves,  # avoid overfit by declining
        'learning_rate': 0.01,
        'min_data_in_leaf': min_data_in_leaf,  # avoid overfit by declining
        'feature_fraction': feature_fraction,
        'bagging_fraction': bagging_fraction,
        'bagging_freq': 5,
        'verbose': 1
    }

    # KFold
    folds = KFold(n_splits=5, random_state=10, shuffle=True)
    CV = np.zeros(len(X_data))
    Test_company = pd.read_excel('testdata/Company.xlsx', encoding='utf-8')
    predictions = np.zeros(len(Test_company))
    train_RMSE = []
    feature_importance_df = pd.DataFrame()
    for fold_, (trn_idx, val_idx) in enumerate(folds.split(X_data, X_label)):

        lgb_train = lgb.Dataset(X_data[trn_idx], label=X_label.iloc[trn_idx])
        lgb_eval = lgb.Dataset(X_data[val_idx], label=X_label.iloc[val_idx])

        gbm = lgb.train(params, lgb_train, num_boost_round=1000, valid_sets=[lgb_train, lgb_eval], verbose_eval=100, early_stopping_rounds=50)

        fold_importance_df = pd.DataFrame()
        fold_importance_df["Feature"] = features
        fold_importance_df["importance"] = gbm.feature_importance()
        fold_importance_df["fold"] = fold_ + 1
        feature_importance_df = pd.concat([feature_importance_df, fold_importance_df], axis=0)

        train_result = gbm.predict(X_data[trn_idx], num_iteration=gbm.best_iteration)
        train_RMSE.append(RMSE(train_result, X_label.iloc[trn_idx].values))

        y_cv = gbm.predict(X_data[val_idx], num_iteration=gbm.best_iteration)
        CV[val_idx] = y_cv

        predictions += gbm.predict(y_data, num_iteration=gbm.best_iteration) / folds.n_splits

    print(train_RMSE)
    RMSE(CV, X_label.values)
    Bias(CV, X_label.values)

    feature_importance_df.to_csv("Fold.csv", index=False, encoding='utf_8_sig')

    for x in range(len(predictions)):
        predictions[x] = round(predictions[x])

    pred_df = pd.DataFrame({"Company": Test_company["Company"].values})
    pred_df["target"] = predictions
    pred_df.to_csv("data/submission.csv", index=False)

    return CV


# SVM
def SVM(X_data, X_label, y_data):
    X_data = X_data.values

    # standarization
    scaler = preprocessing.StandardScaler().fit(X_data)
    X_data = scaler.transform(X_data)
    # y_data = scaler.transform(y_data)

    # PCA, However the result isn't good
    # pca = PCA()  # n_components=20
    # pca.fit(X_data)
    # X_data = pca.transform(X_data)
    # y_data = pca.transform(y_data)

    folds = KFold(n_splits=10, random_state=2319)
    clf = svm.SVR()
    CV = np.zeros(len(X_data))
    for fold_, (trn_idx, val_idx) in enumerate(folds.split(X_data, X_label)):
        print("Fold:%d" % (fold_+1,))
        clf.fit(X_data[trn_idx], X_label.iloc[trn_idx].values.flatten())
        y_cv = clf.predict(X_data[val_idx])
        CV[val_idx] = y_cv

    RMSE(CV, X_label.values)
    Bias(CV, X_label.values)
    return CV


# RandomForest
def RF(X_data, X_label, y_data):
    X_data = X_data.values

    folds = KFold(n_splits=10, random_state=2319)
    regr = RandomForestRegressor(max_depth=2, random_state=12, n_estimators=101)
    CV = np.zeros(len(X_data))
    for fold_, (trn_idx, val_idx) in enumerate(folds.split(X_data, X_label)):
        print("Fold:%d" % (fold_+1,))
        regr.fit(X_data[trn_idx], X_label.iloc[trn_idx].values.flatten())
        y_cv = regr.predict(X_data[val_idx])
        CV[val_idx] = y_cv

    RMSE(CV, X_label.values)
    Bias(CV, X_label.values)
    return CV


# Mix three models to get a better score, it may be better to use a simple model to mix
def Stacking(X_data, X_label, y_data):
    y_cv = []
    y_cv.append(GBDT(X_data, X_label, y_data))
    y_cv.append(SVM(X_data, X_label, y_data))
    y_cv.append(RF(X_data, X_label, y_data))
    y_cv = np.array(y_cv)
    print(y_cv.shape)
    y_cv_median = np.median(y_cv, axis=0)
    print("Median")
    RMSE(y_cv_median, X_label.values)


# Try to get the best parameters
def GridSearch(X_data, X_label, y_data):
    grid = []
    num_leaves = [11, 15, 19, 23]
    min_data_in_leaf = [4, 8, 12, 16]
    feature_fraction = [0.3, 0.5, 0.7, 1]
    bagging_fraction = [0.3, 0.5, 0.7, 1]
    count = 0
    for g1 in num_leaves:
        for g2 in min_data_in_leaf:
            for g3 in feature_fraction:
                for g4 in bagging_fraction:
                    cv = GBDT(X_data, X_label, y_data, num_leaves=g1, min_data_in_leaf=g2, feature_fraction=g3, bagging_fraction=g4)
                    score = RMSE(cv, X_label.values)[0]
                    param = str(g1) + '_' + str(g2) + '_' + str(g3) + '_' + str(g4)
                    grid.append([param, score])
                    count += 1
                    print("Count:%d" % (count))

    with open('GridSearch.csv', 'w', newline='') as f:
        csv_write = csv.writer(f)
        for x in grid:
            csv_write.writerow(x)


# Count RMSE between real and predict
def RMSE(predict, real):
    loss = 0.0
    for index in range(len(predict)):
        loss += (predict[index]-real[index])**2

    print("RMSE:%f" % ((loss/len(predict))**0.5))
    return (loss/len(predict))**0.5


# Count Bias between real and predict
def Bias(predict, real):
    loss = 0.0
    for index in range(len(predict)):
        loss += predict[index]-real[index]

    print('Bias:%f' % (loss/len(predict)))


# Drop the columns with high correlation
def Cor(df):
    print(df.head())
    df_corr = df.corr().values
    high_corr = []
    for x in range(len(df_corr)):
        if x not in high_corr:
            for y in range(x+1, len(df_corr)):
                if np.isnan(df_corr[x][y]):
                    high_corr.append(y)
                elif abs(df_corr[x][y]) >= CORR_LIMIT:
                    high_corr.append(y)
    df.drop(df.columns[high_corr], axis=1, inplace=True)
    print(df.head())
    return df


if __name__ == "__main__":
    Train_data = pd.read_csv('../data/Input.csv', encoding='utf-8')
    Train_data = Train_data.iloc[:, 1:]
    Train_data = Cor(Train_data)
    Train_label = pd.read_csv('../data/Score.csv')
    Train_label = Train_label.iloc[:, 1:]
    Test_data = pd.read_csv('../data/Input.csv', encoding='utf-8')
    Test_data = Test_data.iloc[:, 1:]
    Test_data = Cor(Test_data)

    # model
    GBDT(Train_data, Train_label, Test_data)  # 执行后预测结果会输出至"data/submission.csv"  后续根据比赛要求自行调整成规定格式
    # SVM(Train_data, Train_label, Test_data)
    # RF(Train_data, Train_label, Test_data)

    # Stacking(Train_data, Train_label, Test_data)
    # GridSearch(Train_data, Train_label, Test_data)

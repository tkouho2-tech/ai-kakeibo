import pandas as pd

# Mocking the load_transactions_data logic
values = [
    ["username", "date", "item"],
    ["user1", "2023-01-01", "item1"],
    ["user2", "2023-01-02", "item2"],
    ["yemiko", "2023-01-03", "item3"],
    ["yemiko", "2023-01-04", "item4"],
]

# 1044: records_df = pd.DataFrame(values[1:])
records_df = pd.DataFrame(values[1:], columns=values[0])
print("Original records_df:\n", records_df)

# 1049: records = records_df.to_dict('records')
records = records_df.to_dict('records')

# 1053: df = get_clean_df(records, curr_user)
# Mocking get_clean_df
def get_clean_df_mock(records, username):
    # This is what happens inside get_clean_df:
    # 965: df = pd.DataFrame(records) -> THIS RESETS THE INDEX!
    df = pd.DataFrame(records)
    df = df[df["username"] == username]
    return df

df = get_clean_df_mock(records, "yemiko")
print("\nFiltered df from get_clean_df_mock (Note the reset index if not careful):\n", df)

# 1062: df_all_temp = records_df.copy()
df_all_temp = records_df.copy()
# 1063: df_all_temp['_row_index'] = range(2, len(records) + 2)
df_all_temp['_row_index'] = list(range(2, len(records) + 2))
print("\ndf_all_temp with _row_index:\n", df_all_temp)

# 1067: df = df.join(df_all_temp[['_row_index']])
df_joined = df.join(df_all_temp[['_row_index']])
print("\nJoined df (Final Result):\n", df_joined)

```
# Pick rows where column 'ABCD' starts with 'AV'
av_rows = df[df['ABCD'].str.startswith('AV', na=False)].copy()

# Replace 'MICA Leaf' with value from 'ABCD'
av_rows['MICA Leaf'] = av_rows['ABCD']

# Update Source column
av_rows['Source'] = 'BFA-AV'

# Append back to original dataframe
df = pd.concat([df, av_rows], ignore_index=True)

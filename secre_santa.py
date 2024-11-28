import pandas as pd
import numpy as np
import shutil
import os
import tqdm


def secret_santa(file_path):
    santa_df = pd.read_excel(file_path)
    gift_givers = santa_df['name2'].values
    gift_receivers = santa_df['name2'].copy().values.tolist()
    receiver_interests = santa_df['interests'].values.tolist()
    try:
        shutil.rmtree('output')
    except:
        pass

    os.mkdir('./output/')

    for giver in gift_givers:
        receiver = giver
        counter = 0
        while receiver == giver:
            receiver_index = np.random.randint(0, len(gift_receivers))
            receiver = gift_receivers[receiver_index]
            receiver_interest = receiver_interests[receiver_index].replace(',', ',\n').split(',')
            if counter > 1000:
                try:
                    shutil.rmtree('output')
                except:
                    pass
                return
            counter += 1

        gift_receivers.pop(receiver_index)
        receiver_interests.pop(receiver_index)

        with open(f'output/{giver}.txt', 'w') as f:
            f.write(f'Hellohohohoho {giver}!\n')
            f.write(f'\n')
            f.write(f'You have been chosen to get {receiver.upper()} a gift!\n')
            f.write(f'Their interests are:\n')
            f.write(f'\n')
            f.writelines(receiver_interest)


if __name__ == '__main__':
    n_it = np.random.randint(100)
    for i in tqdm.tqdm(range(n_it)):
        secret_santa('SECRET SECRET SANTA.xlsx')

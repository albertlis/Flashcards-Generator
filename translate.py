import os
import argh
from deep_translator import PonsTranslator
from deep_translator.exceptions import TranslationNotFound
from prodict import Prodict
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor
import threading


class Bar(object):
    def __init__(self, bar):
        self._read_lock = threading.Lock()
        self.bar = bar

    def update_tqdm(self):
        with self._read_lock:
            self.bar.update()


def translate_word(word, counter, source_lang, target_lang, all_translations):
    try:
        translations = PonsTranslator(source=source_lang, target=target_lang).translate(word, return_all=all_translations)
        counter.update_tqdm()
        return Prodict(word=word, translations=translations)
    except TranslationNotFound:
        counter.update_tqdm()
        return Prodict(word=word, translations='Translation not found')


@argh.arg('--source_lang', '-sl')
@argh.arg('--target_lang', '-tl')
@argh.arg('--target_file', '-tf')
@argh.arg('--all_translations', '-all')
def main(input_file, source_lang='english', target_lang='polish', target_file='translations.txt',
         all_translations=False):
    with open(input_file, mode='r', encoding='utf8') as f:
        source_words = [word.strip().lower() for word in f.readlines()]

    with tqdm(source_words) as bar:
        counter = Bar(bar)
        with ThreadPoolExecutor(min(os.cpu_count() * 4, len(source_words))) as pool:
            args = ((source_word, counter, source_lang, target_lang, all_translations) for source_word in source_words)
            translations_info = pool.map(lambda p: translate_word(*p), args)

    text = ''
    for info in translations_info:
        if isinstance(info.translations, list):
            info.translations = ' | '.join(info.translations)
        text += f'{info.translations}\n'

    with open(target_file, mode='w', encoding='utf8') as f:
        f.write(text)


if __name__ == '__main__':
    argh.dispatch_command(main)

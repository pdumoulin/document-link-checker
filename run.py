"""Docx link checker."""

import argparse
import csv
import datetime
import os
import pathlib
import urllib.request

import docx

LOG_LEVELS = [
    'debug',
    'warn',
    'error',
    'none'
]
LOG_LEVEL = None


def main():
    """Entrypoint."""
    args = _load_args()

    # output file preperation
    if os.path.isfile(args.output):
        if not args.f:
            _error(f'output file already exists at {args.output}, use -f to overwrite')  # noqa:E501
    else:
        try:
            pathlib.Path(args.output).touch()
        except Exception as e:
            _error(e)

    # open output file and process data
    with open(args.output, 'w') as csvfile:
        columns = _output_dict(None, None, None).keys()
        writer = csv.DictWriter(csvfile, fieldnames=columns)
        writer.writeheader()

        # generate list of files to examine based on args.target
        files = []
        if os.path.isfile(args.target):
            files.append(args.target)
        elif os.path.isdir(args.target):
            for dirpath, _, filenames in os.walk(args.target):
                for filename in filenames:
                    files.append(os.path.join(dirpath, filename))
        else:
            _error(f'unable to find {args.target} on system')

        # filter out files according to extension
        files = [x for x in files if _allow_file(x)]
        if not files:
            _error(f'Error: unable to find matching file(s) at {args.target} after filtering')  # noqa:E501
        _debug(f'found {len(files)} file(s)')

        # find links in each file
        links = []
        for filename in files:
            links += _extract_links(filename)
        _debug(f'found {len(links)} link(s)')

        # check each link and save results
        error_count = 0
        for filename, url in links:
            error = None
            try:
                req = urllib.request.Request(url)
                req.add_header('User-Agent', args.user_agent)
                urllib.request.urlopen(req, timeout=args.timeout)
            except urllib.error.HTTPError as e:
                error = f'HTTPError: {e.code} response code'
            except urllib.error.URLError as e:
                error = f'URLError: {e.reason}'
            except Exception as e:
                error = f'Unexpected error: {type(e)} | {e}'
            _debug(f'{url} @ {filename}')
            if error:
                writer.writerow(
                    _output_dict(
                        filename,
                        url,
                        error
                    )
                )
                error_count += 1
                _warn(f'{error} @ {url}')

        # output some information
        _debug(f'checked {len(links)} link(s)')
        _debug(f'found {error_count} error(s)')
        _debug(f'saved results to "{args.output}"')


def _debug(message):
    if LOG_LEVELS.index(LOG_LEVEL) <= LOG_LEVELS.index('debug'):
        print('\033[92m' + 'DEBUG:' + '\033[0m', message)


def _warn(message):
    if LOG_LEVELS.index(LOG_LEVEL) <= LOG_LEVELS.index('warn'):
        print('\033[93m' + 'WARN:' + '\033[0m', message)


def _error(message, fatal=True):
    if LOG_LEVELS.index(LOG_LEVEL) <= LOG_LEVELS.index('error'):
        print('\033[91m' + 'ERROR:' + '\033[0m', message)
    if fatal:
        exit(1)


def _output_dict(filename, url, error):
    return {
        'file': filename,
        'url': url,
        'error': error
    }


def _extract_links(filename):
    links = []
    try:
        document = docx.Document(filename)
    except Exception as e:
        _warn(f'unable to open "{filename}" | {type(e)} | {e}')
        return links
    rels = document.part.rels
    for rel_name in rels:
        rel = rels[rel_name]
        if rel.reltype == docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK:  # noqa:E501
            links.append(
                (filename, rel._target)
            )
    return links


def _allow_file(filename, include_suffix=['docx'], exclude_prefix=['.', '~']):
    if any([filename.endswith(x) for x in include_suffix]):
        if all(
            [
                not os.path.split(filename)[-1].startswith(y)
                for y in exclude_prefix
            ]
        ):
            return True
    return False


def _load_args():
    parser = argparse.ArgumentParser(
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument(
        '--target',
        default='.',  # noqa:E501
        help='location of file or directory to load'
    )
    parser.add_argument(
        '--user-agent',
        default='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36',  # noqa:E501
        help='user-agent header used when opening URLs'
    )
    parser.add_argument(
        '--timeout',
        type=int,
        default=10,
        help='time in seconds to give up after when trying to load a URL'
    )
    parser.add_argument(
        '--log-level',
        default=LOG_LEVELS[0],
        choices=LOG_LEVELS,
        help='level of output to STDOUT as script runs'
    )
    today_date = str(datetime.datetime.today().date())
    parser.add_argument(
        '--output',
        default=f'document-check-results-{today_date}.csv',
        help='location of csv file to write results to'
    )
    parser.add_argument(
        '-f',
        default=False,
        action='store_true',
        help='overwrite output file if exists'
    )
    args = parser.parse_args()
    global LOG_LEVEL
    LOG_LEVEL = args.log_level
    return args


if __name__ == '__main__':
    main()

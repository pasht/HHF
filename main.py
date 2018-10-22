import mimetypes

import utils

if __name__ == "__main__":
    print('Start processing....')
    # Initialize mime types
    mimetypes.init();
    mimetypes.add_type('application/x-spss', '.sav')        # mimetype for spss files
    mimetypes.add_type('application/x-stata-dta', '.dta')   # mimetype for stata files

    for entry in utils.walkdir('./test'):
        mimetype = mimetypes.guess_type(entry.path)[0]
        utils.dispatch(entry.path, mimetype)

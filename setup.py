from setuptools import setup

setup(
    name='pdf_merger',
    version='1.0',
    packages=[''],
    install_requires=[
        'PyMuPDF',
        'docx2pdf',
        'PyPDF2',
        'Pillow',
        'pywin32'
    ],
    author='imanuel300',
    author_email='your.email@example.com',
    description='כלי למיזוג קבצי PDF ו-DOCX עם הוספת כותרת תחתונה',
    long_description=open('README.md', encoding='utf-8').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/imanuel300/pdf_merger',
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.7',
)
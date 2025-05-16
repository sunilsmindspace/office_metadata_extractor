from setuptools import setup, find_packages

setup(
    name='office_metadata_extractor',
    version='1.0.0',
    author='Sunil K Sundaram',
    author_email='you@example.com',
    description='Extract custom and core metadata from Office (.docx, .xlsx, .pptx) files',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/sunilsmindspace/office_metadata_extractor',
    packages=find_packages(),
    classifiers=[
        'Programming Language :: Python :: 3',
        'Operating System :: OS Independent',
        'License :: OSI Approved :: MIT',
        'Intended Audience :: Developers',
        'Topic :: File Utilities',
    ],
    python_requires='>=3.13.1',
    include_package_data=True,
    license='MIT'
)

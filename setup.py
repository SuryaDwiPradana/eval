from setuptools import setup, find_packages

with open('README.md') as f:
    readme = f.read()

with open('LICENSE') as f:
    license = f.read()

setup(
    name='org-eval',
    version='0.1.0',
    description='Organization Evaluation',
    long_description=readme,
    author='Surya Dwi Pradana',
    author_email='672016198@student.uksw.edu',
    url='https://github.com/resuryasin/org-eval',
    license=license,
    packages=find_packages(exclude=('tests', 'docs'))
)

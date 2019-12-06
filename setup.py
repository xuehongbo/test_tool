from setuptools import setup, find_packages

setup(
    name='driver_excel-HongBoXue',
    version='0.0.1',
    keywords='driver excel',
    description='a driver excel tool',
    license='MIT License',
    url='https://github.com/xuehongbo/test_tool',
    author='HongBo Xue',
    author_email='505386086@qq.com',
    packages=find_packages(),
    platforms='any',
    install_requires=['openpyxl==3.0.1'],
)

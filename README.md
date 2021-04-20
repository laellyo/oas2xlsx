# oas2xlsx

Welcome to the OAS 2 XSLX home page! 
I developped this console application to facilitate Excel file generation from OAS 2/3 files, that can sometimes be more adapted to analyze API contracts, especially when you have to map fields with external systems. 
Feel free to use and fork it! :)

## Technologies
The project is based on .NET 5 framework.
It is normally multi platforms (Win/MacOS/Linux), but I didn't test it on Linux & MacOS environments.

## Usage

```bash
oas2xlsx -oas <path of your oas file> -xls <output excel file name>
```

:construction: an additional parameter *type* is not yet implemented. Will be used to define the oas source type (file system or url).  

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## Dependencies

This project relies on the following libraries:
- [ClosedXML](https://github.com/ClosedXML/ClosedXML)
- [Microsoft OpenAPI.net](https://github.com/microsoft/OpenAPI.NET)

## License
[MIT](https://choosealicense.com/licenses/mit/)

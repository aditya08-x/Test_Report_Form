from flask_restplus import fields


class NullableString(fields.String):
    """Custom Field for accepting string or null value """
    __schema_type__ = ['string', 'null']
    __schema_example__ = 'string or null'

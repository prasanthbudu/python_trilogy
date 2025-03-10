from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route('/process_data', methods=['POST'])
def process_data():
    try:
        data = request.get_json()
        input_value = data.get('inputData')

        if input_value:
            result = f"Python App Service processed: {input_value}"
            return jsonify({'result': result})
        else:
            return jsonify({'error': 'Missing inputData'}), 400
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=False) #debug should be false in production

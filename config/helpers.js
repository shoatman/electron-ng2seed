var path = require('path');
var _root = path.resolve(__dirname, '..');
function root(args) {
  args = Array.prototype.slice.call(arguments, 0);
  p = path.join.apply(path, [_root].concat(args));
  console.log(p);
  return p;
}
exports.root = root;
